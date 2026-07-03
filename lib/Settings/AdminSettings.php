<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Settings;

use OCA\NeuraMicrosoftSharepointSync\AppInfo\Application;
use OCA\NeuraMicrosoftSharepointSync\Service\ConfigService;
use OCA\NeuraMicrosoftSharepointSync\Service\NextcloudFilesService;
use OCP\AppFramework\Http\TemplateResponse;
use OCP\IURLGenerator;
use OCP\IUserSession;
use OCP\Settings\ISettings;

class AdminSettings implements ISettings {
    public function __construct(
        private ConfigService $configService,
        private NextcloudFilesService $filesService,
        private IURLGenerator $urlGenerator,
        private IUserSession $userSession,
    ) {
    }

    public function getForm(): TemplateResponse {
        $currentUser = $this->userSession->getUser();
        $runAsUser = $this->configService->getRunAsUser();
        if ($runAsUser === '' && $currentUser !== null) {
            $runAsUser = $currentUser->getUID();
        }

        return new TemplateResponse(Application::APP_ID, 'admin', [
            'enabled' => $this->configService->isEnabled(),
            'tenantId' => $this->configService->getTenantId(),
            'clientId' => $this->configService->getClientId(),
            'hasClientSecret' => $this->configService->getClientSecret() !== '',
            'sourceUrl' => $this->configService->getSourceUrl(),
            'destGroupFolderId' => $this->configService->getDestGroupFolderId(),
            'destSubPath' => $this->configService->getDestSubPath(),
            'runAsUser' => $runAsUser,
            'replicateFullPath' => $this->configService->isReplicateFullPath(),
            'connected' => $this->configService->isConnected(),
            'accountName' => $this->configService->getAccountName(),
            'groupFolders' => $this->filesService->listGroupFolders(),
            'redirectUri' => $this->urlGenerator->linkToRouteAbsolute(Application::APP_ID . '.oauth.callback'),
            'oauthStartUrl' => $this->urlGenerator->linkToRoute(Application::APP_ID . '.oauth.start'),
        ], '');
    }

    public function getSection(): string {
        return Application::APP_ID;
    }

    public function getPriority(): int {
        return 50;
    }
}
