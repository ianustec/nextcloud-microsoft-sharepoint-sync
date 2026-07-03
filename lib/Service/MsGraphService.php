<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Service;

use OCP\Http\Client\IClient;
use OCP\Http\Client\IClientService;
use OCP\ITempManager;
use Psr\Log\LoggerInterface;

/**
 * Thin Microsoft Graph client built on top of Nextcloud's HTTP client.
 *
 * Implements the delegated OAuth2 authorization-code flow (single service
 * account) and the handful of Graph endpoints needed to walk and download a
 * shared SharePoint / OneDrive folder tree. No external SDK is required.
 */
class MsGraphService {

    private const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
    private const DOWNLOAD_URL_KEY = '@microsoft.graph.downloadUrl';

    /**
     * Force IPv4 for all outbound requests. Many container networks resolve
     * *.sharepoint.com to IPv6 (AAAA) addresses first but have no working IPv6
     * egress, which makes libcurl fail instantly with "Failed to connect
     * ... after 0 ms". Pinning to IPv4 avoids that dead path.
     */
    private const FORCE_IPV4 = ['force_ip_resolve' => 'v4'];

    public function __construct(
        private ConfigService $configService,
        private IClientService $clientService,
        private ITempManager $tempManager,
        private LoggerInterface $logger,
    ) {
    }

    private function client(): IClient {
        return $this->clientService->newClient();
    }

    private function authorityBase(): string {
        return 'https://login.microsoftonline.com/' . rawurlencode($this->configService->getTenantId()) . '/oauth2/v2.0';
    }

    // ---------------------------------------------------------------------
    // OAuth2 authorization-code flow
    // ---------------------------------------------------------------------

    /** Builds the Microsoft sign-in URL the admin is redirected to. */
    public function getAuthorizeUrl(string $state, string $redirectUri): string {
        $params = [
            'client_id' => $this->configService->getClientId(),
            'response_type' => 'code',
            'redirect_uri' => $redirectUri,
            'response_mode' => 'query',
            'scope' => ConfigService::SCOPES,
            'state' => $state,
            'prompt' => 'select_account',
        ];
        return $this->authorityBase() . '/authorize?' . http_build_query($params);
    }

    /**
     * Exchanges an authorization code for tokens, stores the refresh token and
     * the connected account display name.
     *
     * @throws \RuntimeException on failure.
     */
    public function exchangeCode(string $code, string $redirectUri): void {
        $data = $this->tokenRequest([
            'client_id' => $this->configService->getClientId(),
            'client_secret' => $this->configService->getClientSecret(),
            'scope' => ConfigService::SCOPES,
            'code' => $code,
            'redirect_uri' => $redirectUri,
            'grant_type' => 'authorization_code',
        ]);

        if (empty($data['refresh_token'])) {
            throw new \RuntimeException('Microsoft did not return a refresh token. Ensure "offline_access" is granted.');
        }

        $this->configService->setRefreshToken((string)$data['refresh_token']);
        if (!empty($data['access_token'])) {
            $this->configService->cacheAccessToken((string)$data['access_token'], (int)($data['expires_in'] ?? 3600));
        }

        try {
            $this->configService->setAccountName($this->fetchAccountName((string)$data['access_token']));
        } catch (\Throwable $e) {
            $this->logger->warning('Could not fetch account name: ' . $e->getMessage(), ['app' => 'neura_microsoft_sharepoint_sync']);
        }
    }

    /**
     * Returns a valid access token, using the cached one when still valid and
     * refreshing via the stored refresh token otherwise.
     *
     * @throws SharepointSyncSkipException when the app is not connected.
     * @throws \RuntimeException on refresh failure.
     */
    public function getAccessToken(): string {
        $cached = $this->configService->getCachedAccessToken();
        if ($cached !== '' && $this->configService->getAccessTokenExpiry() > time() + 60) {
            return $cached;
        }

        $refreshToken = $this->configService->getRefreshToken();
        if ($refreshToken === '') {
            throw new SharepointSyncSkipException('Not connected to Microsoft. Click "Connect Microsoft" first.');
        }

        try {
            $data = $this->tokenRequest([
                'client_id' => $this->configService->getClientId(),
                'client_secret' => $this->configService->getClientSecret(),
                'scope' => ConfigService::SCOPES,
                'refresh_token' => $refreshToken,
                'grant_type' => 'refresh_token',
            ]);
        } catch (\Throwable $e) {
            // A failed refresh usually means the grant was revoked/expired.
            if (stripos($e->getMessage(), 'invalid_grant') !== false) {
                $this->configService->clearTokens();
            }
            throw new \RuntimeException('Token refresh failed: ' . $e->getMessage(), 0, $e);
        }

        $accessToken = (string)($data['access_token'] ?? '');
        if ($accessToken === '') {
            throw new \RuntimeException('Token refresh returned no access token.');
        }
        // Microsoft rotates refresh tokens; persist the new one when provided.
        if (!empty($data['refresh_token'])) {
            $this->configService->setRefreshToken((string)$data['refresh_token']);
        }
        $this->configService->cacheAccessToken($accessToken, (int)($data['expires_in'] ?? 3600));
        return $accessToken;
    }

    /**
     * @param array<string, string> $params
     * @return array<string, mixed>
     */
    private function tokenRequest(array $params): array {
        try {
            $response = $this->client()->post($this->authorityBase() . '/token', self::FORCE_IPV4 + [
                'body' => http_build_query($params),
                'headers' => ['Content-Type' => 'application/x-www-form-urlencoded'],
                'timeout' => 30,
            ]);
        } catch (\Throwable $e) {
            throw new \RuntimeException($this->extractError($e), 0, $e);
        }
        $decoded = json_decode((string)$response->getBody(), true);
        if (!is_array($decoded)) {
            throw new \RuntimeException('Invalid token endpoint response.');
        }
        if (isset($decoded['error'])) {
            throw new \RuntimeException($decoded['error'] . ': ' . ($decoded['error_description'] ?? ''));
        }
        return $decoded;
    }

    // ---------------------------------------------------------------------
    // Graph API calls
    // ---------------------------------------------------------------------

    /** Returns the display name / UPN of the connected account. */
    public function fetchAccountName(string $accessToken): string {
        $me = $this->graphGet(self::GRAPH_BASE . '/me?$select=displayName,userPrincipalName,mail', $accessToken);
        return (string)($me['displayName'] ?? '') !== ''
            ? (string)$me['displayName'] . ' <' . (string)($me['userPrincipalName'] ?? $me['mail'] ?? '') . '>'
            : (string)($me['userPrincipalName'] ?? $me['mail'] ?? 'Microsoft account');
    }

    /** Convenience wrapper used by the admin "Test connection" button. */
    public function getConnectedAccountName(): string {
        return $this->fetchAccountName($this->getAccessToken());
    }

    /**
     * Resolves a SharePoint / OneDrive folder URL (or sharing link) to the
     * DriveItem it points to.
     *
     * @return array{driveId: string, itemId: string, name: string, relativePath: string}
     * @throws \RuntimeException when the share cannot be resolved.
     */
    public function resolveShare(string $url): array {
        $token = $this->getAccessToken();
        $encoded = $this->encodeSharingUrl($url);
        $endpoint = self::GRAPH_BASE . '/shares/' . $encoded . '/driveItem?$select=id,name,parentReference,folder,file';
        $item = $this->graphGet($endpoint, $token);

        $itemId = (string)($item['id'] ?? '');
        $driveId = (string)($item['parentReference']['driveId'] ?? '');
        if ($itemId === '' || $driveId === '') {
            throw new \RuntimeException('Could not resolve the SharePoint URL to a drive item.');
        }
        if (!isset($item['folder'])) {
            throw new \RuntimeException('The SharePoint URL points to a file, not a folder.');
        }

        return [
            'driveId' => $driveId,
            'itemId' => $itemId,
            'name' => (string)($item['name'] ?? 'SharePoint'),
            'relativePath' => $this->relativePathFromItem($item),
        ];
    }

    /**
     * Lists the direct children of a drive item, following pagination.
     *
     * @return list<array<string, mixed>>
     */
    public function listChildren(string $driveId, string $itemId): array {
        $token = $this->getAccessToken();
        $url = self::GRAPH_BASE . '/drives/' . rawurlencode($driveId) . '/items/' . rawurlencode($itemId) . '/children?$top=200';
        $items = [];
        while ($url !== null) {
            $page = $this->graphGet($url, $token);
            foreach (($page['value'] ?? []) as $child) {
                $items[] = $child;
            }
            $url = isset($page['@odata.nextLink']) ? (string)$page['@odata.nextLink'] : null;
        }
        return $items;
    }

    /**
     * Downloads a drive item's content to a temporary file and returns its path.
     *
     * Strategy:
     *   1. Try the pre-signed "@microsoft.graph.downloadUrl" embedded in the item
     *      (no auth header needed, fastest path).
     *   2. If that URL returns an HTML page (it has expired – pre-signed URLs
     *      typically live ~1 h and large folder trees take longer to traverse),
     *      fall back to the Graph API /content endpoint with a freshly-fetched
     *      Bearer token.
     *
     * @param array<string, mixed> $item A child item as returned by listChildren().
     * @throws \RuntimeException on failure.
     */
    public function downloadToTemp(string $driveId, array $item): string {
        $tmp = $this->tempManager->getTemporaryFile();
        $downloadUrl = $item[self::DOWNLOAD_URL_KEY] ?? null;
        $expectedSize = (int)($item['size'] ?? -1);

        try {
            $presignedOk = false;
            if (is_string($downloadUrl) && $downloadUrl !== '') {
                $response = $this->client()->get($downloadUrl, self::FORCE_IPV4 + ['sink' => $tmp, 'timeout' => 600]);
                $this->honourSinkFallback($response, $tmp, $expectedSize);

                if ($this->isHtmlContent($tmp)) {
                    // Pre-signed URL returned an HTML error/login page (expired).
                    // Wipe the temp file and fall through to the Graph API path.
                    file_put_contents($tmp, '');
                    $this->logger->debug(
                        'Pre-signed download URL returned HTML for item "' . ($item['name'] ?? '') . '"; retrying via Graph API.',
                        ['app' => 'neura_microsoft_sharepoint_sync']
                    );
                } else {
                    $presignedOk = true;
                }
            }

            if (!$presignedOk) {
                $token = $this->getAccessToken();
                $url = self::GRAPH_BASE . '/drives/' . rawurlencode($driveId) . '/items/' . rawurlencode((string)$item['id']) . '/content';
                $response = $this->client()->get($url, self::FORCE_IPV4 + [
                    'sink' => $tmp,
                    'timeout' => 600,
                    'headers' => ['Authorization' => 'Bearer ' . $token],
                ]);
                $this->honourSinkFallback($response, $tmp, $expectedSize);
            }
        } catch (\Throwable $e) {
            throw new \RuntimeException($this->extractError($e), 0, $e);
        }
        return $tmp;
    }

    /**
     * If the HTTP client did not honour the "sink" option (i.e. the file is
     * still empty after the request), write the response body ourselves.
     */
    private function honourSinkFallback(mixed $response, string $tmp, int $expectedSize): void {
        if ($expectedSize > 0 && (int)@filesize($tmp) === 0) {
            $body = $response->getBody();
            if (is_resource($body)) {
                $dest = fopen($tmp, 'wb');
                if ($dest !== false) {
                    stream_copy_to_stream($body, $dest);
                    fclose($dest);
                }
            } elseif (is_string($body) && $body !== '') {
                file_put_contents($tmp, $body);
            }
        }
    }

    /**
     * Returns true when the first bytes of a downloaded temp file look like an
     * HTML page rather than binary file content. SharePoint redirects expired
     * pre-signed URLs to a login/error HTML page.
     */
    private function isHtmlContent(string $path): bool {
        $handle = @fopen($path, 'rb');
        if ($handle === false) {
            return false;
        }
        $start = (string)fread($handle, 20);
        fclose($handle);
        return stripos(ltrim($start), '<!DOCTYPE') === 0 || stripos(ltrim($start), '<html') === 0;
    }

    // ---------------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------------

    /**
     * @return array<string, mixed>
     * @throws \RuntimeException
     */
    private function graphGet(string $url, string $accessToken): array {
        try {
            $response = $this->client()->get($url, self::FORCE_IPV4 + [
                'headers' => [
                    'Authorization' => 'Bearer ' . $accessToken,
                    'Accept' => 'application/json',
                ],
                'timeout' => 60,
            ]);
        } catch (\Throwable $e) {
            throw new \RuntimeException($this->extractError($e), 0, $e);
        }
        $decoded = json_decode((string)$response->getBody(), true);
        if (!is_array($decoded)) {
            throw new \RuntimeException('Invalid Graph response.');
        }
        if (isset($decoded['error'])) {
            $err = $decoded['error'];
            $msg = is_array($err) ? ($err['message'] ?? $err['code'] ?? 'Graph error') : (string)$err;
            throw new \RuntimeException((string)$msg);
        }
        return $decoded;
    }

    /** Encodes a sharing URL into the shareId format expected by /shares. */
    private function encodeSharingUrl(string $url): string {
        $base64 = base64_encode($url);
        return 'u!' . rtrim(strtr($base64, '+/', '-_'), '=');
    }

    /**
     * Derives the folder path relative to the drive root from a DriveItem's
     * parentReference.path + name, e.g. "Campaigns/2026/Q1-Report".
     *
     * @param array<string, mixed> $item
     */
    private function relativePathFromItem(array $item): string {
        $parentPath = (string)($item['parentReference']['path'] ?? '');
        $name = (string)($item['name'] ?? '');
        // parentPath looks like "/drives/{id}/root:/A/B" or "/drive/root:/A/B".
        $rel = '';
        if (($pos = strpos($parentPath, 'root:')) !== false) {
            $rel = substr($parentPath, $pos + strlen('root:'));
        }
        $rel = rawurldecode(trim($rel, '/'));
        $segments = array_filter(array_map('trim', $rel === '' ? [] : explode('/', $rel)), static fn ($s) => $s !== '');
        $segments[] = $name;
        return implode('/', $segments);
    }

    /** Best-effort extraction of a human-readable error from an HTTP exception. */
    private function extractError(\Throwable $e): string {
        $message = $e->getMessage();
        // Guzzle exceptions expose the response with the JSON error body.
        if (method_exists($e, 'getResponse')) {
            /** @var mixed $resp */
            $resp = $e->getResponse();
            if ($resp !== null && method_exists($resp, 'getBody')) {
                $body = (string)$resp->getBody();
                $decoded = json_decode($body, true);
                if (isset($decoded['error'])) {
                    $err = $decoded['error'];
                    if (is_array($err)) {
                        return (string)($err['message'] ?? $err['code'] ?? $message);
                    }
                    return (string)$err . ': ' . (string)($decoded['error_description'] ?? '');
                }
                if ($body !== '') {
                    return substr($body, 0, 500);
                }
            }
        }
        return $message;
    }
}
