<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Service;

/**
 * Non-fatal signal used to abort a sync run for an expected, recoverable reason
 * (missing configuration, not connected, share not accessible). The admin UI
 * reports these as "skipped" rather than hard errors.
 */
class SharepointSyncSkipException extends \RuntimeException {
}
