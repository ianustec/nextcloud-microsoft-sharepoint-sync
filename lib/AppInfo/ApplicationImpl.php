<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\AppInfo;

use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

/**
 * Application entry point for the Microsoft SharePoint File Sync app.
 *
 * Settings are declared in appinfo/info.xml. All services rely on
 * Nextcloud's built-in HTTP client, so there are no external Composer
 * dependencies to load at boot time.
 */
class Application extends App implements IBootstrap {

    public const APP_ID = 'neura_microsoft_sharepoint_sync';

    public function __construct() {
        parent::__construct(self::APP_ID);
    }

    public function register(IRegistrationContext $context): void {
    }

    public function boot(IBootContext $context): void {
    }
}
