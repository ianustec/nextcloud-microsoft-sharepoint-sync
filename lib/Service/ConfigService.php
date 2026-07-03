<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Service;

use OCA\NeuraMicrosoftSharepointSync\AppInfo\Application;
use OCP\IConfig;
use OCP\Security\ICrypto;

/**
 * Centralises all app configuration stored via Nextcloud's IConfig.
 *
 * Sensitive values (client secret, refresh token, cached access token) are
 * encrypted at rest using Nextcloud's ICrypto (AES-256-GCM). A fallback to the
 * raw stored value is provided so that decryption failures never crash the app.
 */
class ConfigService {

    // Azure AD app registration
    public const KEY_TENANT_ID     = 'tenant_id';
    public const KEY_CLIENT_ID     = 'client_id';
    public const KEY_CLIENT_SECRET = 'client_secret';

    // OAuth tokens (delegated flow, single service account)
    public const KEY_REFRESH_TOKEN = 'refresh_token';
    public const KEY_ACCESS_TOKEN  = 'access_token';
    public const KEY_TOKEN_EXPIRY  = 'access_token_expiry';
    public const KEY_ACCOUNT_NAME  = 'account_name';

    // Source / destination
    public const KEY_SOURCE_URL      = 'source_url';
    public const KEY_DEST_GF_ID      = 'dest_groupfolder_id';
    public const KEY_DEST_SUBPATH    = 'dest_subpath';
    public const KEY_RUNAS_USER      = 'runas_user';
    public const KEY_REPLICATE_PATH  = 'replicate_full_path';
    public const KEY_ENABLED         = 'enabled';

    /** Delegated Microsoft Graph scopes required by the app. */
    public const SCOPES = 'offline_access openid profile User.Read Files.Read.All Sites.Read.All';

    public function __construct(
        private IConfig $config,
        private ICrypto $crypto,
    ) {
    }

    private function get(string $key, string $default = ''): string {
        return $this->config->getAppValue(Application::APP_ID, $key, $default);
    }

    private function set(string $key, string $value): void {
        $this->config->setAppValue(Application::APP_ID, $key, $value);
    }

    private function getEncrypted(string $key): string {
        $stored = $this->get($key);
        if ($stored === '') {
            return '';
        }
        try {
            return $this->crypto->decrypt($stored);
        } catch (\Throwable) {
            // Value was stored before encryption or is corrupt; return as-is.
            return $stored;
        }
    }

    private function setEncrypted(string $key, string $value): void {
        if ($value === '') {
            $this->config->deleteAppValue(Application::APP_ID, $key);
            return;
        }
        $this->set($key, $this->crypto->encrypt($value));
    }

    // ---- Azure AD app registration ----

    /** Tenant id or "common"/"organizations". Defaults to "common". */
    public function getTenantId(): string {
        $t = trim($this->get(self::KEY_TENANT_ID));
        return $t !== '' ? $t : 'common';
    }

    public function setTenantId(string $tenantId): void {
        $this->set(self::KEY_TENANT_ID, trim($tenantId));
    }

    public function getClientId(): string {
        return trim($this->get(self::KEY_CLIENT_ID));
    }

    public function setClientId(string $clientId): void {
        $this->set(self::KEY_CLIENT_ID, trim($clientId));
    }

    public function getClientSecret(): string {
        return $this->getEncrypted(self::KEY_CLIENT_SECRET);
    }

    public function setClientSecret(string $secret): void {
        $this->setEncrypted(self::KEY_CLIENT_SECRET, trim($secret));
    }

    public function hasAppCredentials(): bool {
        return $this->getClientId() !== '' && $this->getClientSecret() !== '';
    }

    // ---- OAuth tokens ----

    public function getRefreshToken(): string {
        return $this->getEncrypted(self::KEY_REFRESH_TOKEN);
    }

    public function setRefreshToken(string $token): void {
        $this->setEncrypted(self::KEY_REFRESH_TOKEN, $token);
    }

    public function isConnected(): bool {
        return $this->getRefreshToken() !== '';
    }

    public function getCachedAccessToken(): string {
        return $this->getEncrypted(self::KEY_ACCESS_TOKEN);
    }

    public function getAccessTokenExpiry(): int {
        return (int)$this->get(self::KEY_TOKEN_EXPIRY, '0');
    }

    /** Stores the access token together with its absolute expiry timestamp. */
    public function cacheAccessToken(string $token, int $expiresInSeconds): void {
        $this->setEncrypted(self::KEY_ACCESS_TOKEN, $token);
        $this->set(self::KEY_TOKEN_EXPIRY, (string)(time() + max(0, $expiresInSeconds)));
    }

    public function getAccountName(): string {
        return $this->get(self::KEY_ACCOUNT_NAME);
    }

    public function setAccountName(string $name): void {
        $this->set(self::KEY_ACCOUNT_NAME, $name);
    }

    /** Removes all stored tokens and account info (disconnect). */
    public function clearTokens(): void {
        $this->config->deleteAppValue(Application::APP_ID, self::KEY_REFRESH_TOKEN);
        $this->config->deleteAppValue(Application::APP_ID, self::KEY_ACCESS_TOKEN);
        $this->config->deleteAppValue(Application::APP_ID, self::KEY_TOKEN_EXPIRY);
        $this->config->deleteAppValue(Application::APP_ID, self::KEY_ACCOUNT_NAME);
    }

    // ---- Source / destination ----

    public function getSourceUrl(): string {
        return trim($this->get(self::KEY_SOURCE_URL));
    }

    public function setSourceUrl(string $url): void {
        $this->set(self::KEY_SOURCE_URL, trim($url));
    }

    public function getDestGroupFolderId(): int {
        return (int)$this->get(self::KEY_DEST_GF_ID, '0');
    }

    public function setDestGroupFolderId(int $id): void {
        $this->set(self::KEY_DEST_GF_ID, (string)$id);
    }

    /** Optional sub-path inside the group folder where the tree is created. */
    public function getDestSubPath(): string {
        return trim($this->get(self::KEY_DEST_SUBPATH), " \t\n\r\0\x0B/");
    }

    public function setDestSubPath(string $path): void {
        $this->set(self::KEY_DEST_SUBPATH, trim($path, " \t\n\r\0\x0B/"));
    }

    /** User whose filesystem is used to write into the group folder mount. */
    public function getRunAsUser(): string {
        return trim($this->get(self::KEY_RUNAS_USER));
    }

    public function setRunAsUser(string $userId): void {
        $this->set(self::KEY_RUNAS_USER, trim($userId));
    }

    /**
     * Whether to replicate the full SharePoint path (relative to the drive root)
     * under the destination, or only the leaf folder name. Defaults to true.
     */
    public function isReplicateFullPath(): bool {
        return $this->get(self::KEY_REPLICATE_PATH, 'yes') === 'yes';
    }

    public function setReplicateFullPath(bool $v): void {
        $this->set(self::KEY_REPLICATE_PATH, $v ? 'yes' : 'no');
    }

    public function isEnabled(): bool {
        return $this->get(self::KEY_ENABLED, 'yes') === 'yes';
    }

    public function setEnabled(bool $enabled): void {
        $this->set(self::KEY_ENABLED, $enabled ? 'yes' : 'no');
    }
}
