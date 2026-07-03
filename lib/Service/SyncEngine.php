<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Service;

use OCP\Files\Folder;
use Psr\Log\LoggerInterface;

/**
 * Orchestrates a one-way, recursive download from a SharePoint / OneDrive
 * folder into a Nextcloud Group Folder.
 *
 * The SharePoint folder hierarchy is replicated idempotently: existing folders
 * are reused and only missing ones are created. Files are (re)downloaded only
 * when their size or modification time differs from the copy already stored.
 */
class SyncEngine {

    public function __construct(
        private ConfigService $configService,
        private MsGraphService $graphService,
        private NextcloudFilesService $filesService,
        private LoggerInterface $logger,
    ) {
    }

    /**
     * Runs the full recursive download.
     *
     * @return array{folders: int, downloaded: int, skipped: int, failed: int, errors: list<string>, root: string}
     * @throws SharepointSyncSkipException when preconditions are not met.
     * @throws \RuntimeException on fatal errors.
     */
    public function run(): array {
        if (!$this->configService->isEnabled()) {
            throw new SharepointSyncSkipException('Sync is disabled in the settings.');
        }
        if (!$this->configService->isConnected()) {
            throw new SharepointSyncSkipException('Not connected to Microsoft. Click "Connect Microsoft" first.');
        }
        $sourceUrl = $this->configService->getSourceUrl();
        if ($sourceUrl === '') {
            throw new SharepointSyncSkipException('No SharePoint source URL configured.');
        }
        if ($this->configService->getDestGroupFolderId() <= 0) {
            throw new SharepointSyncSkipException('No destination group folder selected.');
        }

        $report = ['folders' => 0, 'downloaded' => 0, 'skipped' => 0, 'failed' => 0, 'errors' => [], 'root' => ''];

        $root = $this->graphService->resolveShare($sourceUrl);

        $base = $this->filesService->getDestinationBase(
            $this->configService->getDestGroupFolderId(),
            $this->configService->getDestSubPath(),
            $this->configService->getRunAsUser(),
        );

        $relative = $this->configService->isReplicateFullPath() ? $root['relativePath'] : $root['name'];
        $targetRoot = $this->filesService->ensureFolderPath($base, $relative);
        $report['root'] = $relative;

        $this->logger->info(
            'SharePoint sync started: "' . $sourceUrl . '" -> group folder #'
            . $this->configService->getDestGroupFolderId() . '/' . $relative,
            ['app' => 'neura_microsoft_sharepoint_sync']
        );

        // Iterative DFS to avoid deep recursion on large trees.
        $stack = [[$root['driveId'], $root['itemId'], $targetRoot]];
        while ($stack !== []) {
            /** @var array{0: string, 1: string, 2: Folder} $frame */
            $frame = array_pop($stack);
            [$driveId, $itemId, $ncFolder] = $frame;
            $report['folders']++;

            try {
                $children = $this->graphService->listChildren($driveId, $itemId);
            } catch (\Throwable $e) {
                $report['failed']++;
                $report['errors'][] = 'List failed for "' . $ncFolder->getName() . '": ' . $e->getMessage();
                continue;
            }

            foreach ($children as $child) {
                $name = (string)($child['name'] ?? '');
                if ($name === '') {
                    continue;
                }

                if (isset($child['folder'])) {
                    try {
                        $sub = $this->filesService->ensureFolderPath($ncFolder, $name);
                        $stack[] = [$driveId, (string)$child['id'], $sub];
                    } catch (\Throwable $e) {
                        $report['failed']++;
                        $report['errors'][] = 'Folder "' . $name . '": ' . $e->getMessage();
                    }
                    continue;
                }

                try {
                    $this->syncFile($driveId, $child, $ncFolder, $report);
                } catch (\Throwable $e) {
                    $report['failed']++;
                    $report['errors'][] = 'File "' . $name . '": ' . $e->getMessage();
                }
            }
        }

        $this->logger->info(
            'SharePoint sync finished: ' . $report['downloaded'] . ' downloaded, '
            . $report['skipped'] . ' skipped, ' . $report['failed'] . ' failed.',
            ['app' => 'neura_microsoft_sharepoint_sync']
        );

        return $report;
    }

    /**
     * Downloads a single file into $ncFolder unless an up-to-date copy exists.
     *
     * @param array<string, mixed> $child
     * @param array{folders: int, downloaded: int, skipped: int, failed: int, errors: list<string>, root: string} $report
     */
    private function syncFile(string $driveId, array $child, Folder $ncFolder, array &$report): void {
        $name = (string)$child['name'];
        $spSize = (int)($child['size'] ?? -1);
        $spMtime = $this->parseTimestamp((string)($child['lastModifiedDateTime'] ?? ''));

        $existing = $this->filesService->getExistingFileInfo($ncFolder, $name);
        if ($existing !== null && $spSize >= 0 && $existing['size'] === $spSize && $existing['mtime'] >= $spMtime && $spMtime > 0) {
            $report['skipped']++;
            return;
        }

        $tmp = $this->graphService->downloadToTemp($driveId, $child);
        try {
            $this->filesService->putFileFromPath($ncFolder, $name, $tmp, $spMtime);
            $report['downloaded']++;
        } finally {
            if (is_file($tmp)) {
                @unlink($tmp);
            }
        }
    }

    private function parseTimestamp(string $iso): int {
        if ($iso === '') {
            return 0;
        }
        try {
            return (new \DateTimeImmutable($iso))->getTimestamp();
        } catch (\Throwable) {
            return 0;
        }
    }
}
