<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Controller;

use OCA\NeuraMicrosoftSharepointSync\AppInfo\Application;
use OCA\NeuraMicrosoftSharepointSync\Service\ConfigService;
use OCA\NeuraMicrosoftSharepointSync\Service\MsGraphService;
use OCA\NeuraMicrosoftSharepointSync\Service\NextcloudFilesService;
use OCA\NeuraMicrosoftSharepointSync\Service\SharepointSyncSkipException;
use OCA\NeuraMicrosoftSharepointSync\Service\SyncEngine;
use OCP\AppFramework\Controller;
use OCP\AppFramework\Http\DataResponse;
use OCP\IRequest;

/**
 * HTTP controller for the admin settings panel.
 *
 * All endpoints require admin privileges (enforced by @AdminRequired).
 */
class AdminSettingsController extends Controller {

    public function __construct(
        IRequest $request,
        private ConfigService $configService,
        private MsGraphService $graphService,
        private NextcloudFilesService $filesService,
        private SyncEngine $syncEngine,
    ) {
        parent::__construct(Application::APP_ID, $request);
    }

    /**
     * Persists the admin settings. The client secret is only updated when a
     * non-empty value is provided, so reloading the form without retyping it
     * does not erase the stored secret.
     *
     * @AdminRequired
     */
    public function save(
        string $tenantId = 'common',
        string $clientId = '',
        string $clientSecret = '',
        string $sourceUrl = '',
        int $destGroupFolderId = 0,
        string $destSubPath = '',
        string $runAsUser = '',
        string $replicateFullPath = 'yes',
        string $enabled = 'yes',
    ): DataResponse {
        $this->configService->setTenantId($tenantId);
        $this->configService->setClientId($clientId);
        if ($clientSecret !== '') {
            $this->configService->setClientSecret($clientSecret);
        }
        $this->configService->setSourceUrl($sourceUrl);
        $this->configService->setDestGroupFolderId($destGroupFolderId);
        $this->configService->setDestSubPath($destSubPath);
        $this->configService->setRunAsUser($runAsUser);
        $this->configService->setReplicateFullPath($replicateFullPath === 'yes');
        $this->configService->setEnabled($enabled === 'yes');

        return new DataResponse(['status' => 'ok']);
    }

    /**
     * Verifies the Microsoft connection by fetching the connected account.
     *
     * @AdminRequired
     */
    public function testConnection(): DataResponse {
        if (!$this->configService->isConnected()) {
            return new DataResponse(['status' => 'error', 'message' => 'Not connected. Click "Connect Microsoft" first.'], 400);
        }
        try {
            $account = $this->graphService->getConnectedAccountName();
            return new DataResponse(['status' => 'ok', 'account' => $account]);
        } catch (\Throwable $e) {
            return new DataResponse(['status' => 'error', 'message' => $e->getMessage()], 400);
        }
    }

    /**
     * Returns the list of Group Folders available as sync destinations.
     *
     * @AdminRequired
     */
    public function listGroupFolders(): DataResponse {
        return new DataResponse(['status' => 'ok', 'groupFolders' => $this->filesService->listGroupFolders()]);
    }

    /**
     * Removes the stored Microsoft tokens (disconnect the account).
     *
     * @AdminRequired
     */
    public function disconnect(): DataResponse {
        $this->configService->clearTokens();
        return new DataResponse(['status' => 'ok']);
    }

    /**
     * Runs the recursive SharePoint download and returns a summary report.
     *
     * Source and destination are taken live from the form so the operator can
     * change the folder to import without re-saving the connection settings.
     * The Microsoft credentials (tenant/client/secret) are saved separately.
     *
     * @AdminRequired
     */
    public function syncNow(
        string $sourceUrl = '',
        int $destGroupFolderId = 0,
        string $destSubPath = '',
        string $runAsUser = '',
        string $replicateFullPath = 'yes',
    ): DataResponse {
        set_time_limit(3600);
        @ini_set('max_execution_time', '3600');

        // Persist the source/destination in real time (no separate "Save" needed).
        if ($sourceUrl !== '') {
            $this->configService->setSourceUrl($sourceUrl);
        }
        if ($destGroupFolderId > 0) {
            $this->configService->setDestGroupFolderId($destGroupFolderId);
        }
        $this->configService->setDestSubPath($destSubPath);
        if ($runAsUser !== '') {
            $this->configService->setRunAsUser($runAsUser);
        }
        $this->configService->setReplicateFullPath($replicateFullPath === 'yes');

        try {
            $report = $this->syncEngine->run();
            return new DataResponse(['status' => 'ok'] + $report);
        } catch (SharepointSyncSkipException $e) {
            return new DataResponse(['status' => 'skipped', 'message' => $e->getMessage()], 200);
        } catch (\Throwable $e) {
            return new DataResponse(['status' => 'error', 'message' => $e->getMessage()], 200);
        }
    }
}
