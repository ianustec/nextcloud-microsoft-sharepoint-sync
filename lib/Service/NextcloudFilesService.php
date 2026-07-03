<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Service;

use OCP\DB\QueryBuilder\IQueryBuilder;
use OCP\Files\Folder;
use OCP\Files\IRootFolder;
use OCP\Files\Node;
use OCP\Files\NotFoundException;
use OCP\IDBConnection;
use Psr\Log\LoggerInterface;

/**
 * Nextcloud-side adapter: lists Group Folders and writes downloaded files into
 * them, replicating the SharePoint folder tree idempotently.
 *
 * Group Folders are read directly from the "group_folders" table (stable across
 * versions) and written through a member user's filesystem view, so no hard
 * dependency on the groupfolders app classes is required.
 */
class NextcloudFilesService {

    public function __construct(
        private IRootFolder $rootFolder,
        private IDBConnection $db,
        private LoggerInterface $logger,
    ) {
    }

    /**
     * Returns the configured Group Folders as [{id, mountPoint}], or an empty
     * list when the groupfolders app is not installed.
     *
     * @return list<array{id: int, mountPoint: string}>
     */
    public function listGroupFolders(): array {
        try {
            $qb = $this->db->getQueryBuilder();
            $qb->select('folder_id', 'mount_point')
                ->from('group_folders')
                ->orderBy('mount_point', 'ASC');
            $result = $qb->executeQuery();
            $folders = [];
            while ($row = $result->fetch()) {
                $folders[] = [
                    'id' => (int)$row['folder_id'],
                    'mountPoint' => (string)$row['mount_point'],
                ];
            }
            $result->closeCursor();
            return $folders;
        } catch (\Throwable $e) {
            $this->logger->info('Could not list group folders (app not installed?): ' . $e->getMessage(), ['app' => 'neura_microsoft_sharepoint_sync']);
            return [];
        }
    }

    public function getGroupFolderMountPoint(int $folderId): ?string {
        foreach ($this->listGroupFolders() as $gf) {
            if ($gf['id'] === $folderId) {
                return $gf['mountPoint'];
            }
        }
        return null;
    }

    /**
     * Resolves the destination base folder inside the chosen Group Folder,
     * creating the optional sub-path if needed.
     *
     * @throws \RuntimeException when the group folder is not reachable for the user.
     */
    public function getDestinationBase(int $groupFolderId, string $subPath, string $runAsUser): Folder {
        if ($runAsUser === '') {
            throw new \RuntimeException('No "run as" user configured for writing into the group folder.');
        }
        $mountPoint = $this->getGroupFolderMountPoint($groupFolderId);
        if ($mountPoint === null) {
            throw new \RuntimeException('Group folder #' . $groupFolderId . ' not found. Is the Group Folders app enabled?');
        }

        $userFolder = $this->rootFolder->getUserFolder($runAsUser);
        try {
            $base = $userFolder->get($mountPoint);
        } catch (NotFoundException) {
            throw new \RuntimeException(
                'User "' . $runAsUser . '" cannot access group folder "' . $mountPoint . '". '
                . 'Add this user to a group that has write access to the folder.'
            );
        }
        if (!$base instanceof Folder) {
            throw new \RuntimeException('Group folder mount "' . $mountPoint . '" is not a folder.');
        }

        return $subPath === '' ? $base : $this->ensureFolderPath($base, $subPath);
    }

    /**
     * Ensures a (possibly nested) relative path exists under $parent, reusing
     * existing folders and only creating the missing segments.
     */
    public function ensureFolderPath(Folder $parent, string $relPath): Folder {
        $current = $parent;
        foreach ($this->splitPath($relPath) as $segment) {
            $safe = $this->sanitizeName($segment);
            try {
                $node = $current->get($safe);
                if (!$node instanceof Folder) {
                    throw new \RuntimeException('Expected a folder at "' . $safe . '" but found a file.');
                }
                $current = $node;
            } catch (NotFoundException) {
                $current = $current->newFolder($safe);
            }
        }
        return $current;
    }

    /**
     * Writes a file into $parent from a local temp path, overwriting when it
     * already exists. Preserves the SharePoint modification time.
     */
    public function putFileFromPath(Folder $parent, string $name, string $tmpPath, int $mtime = 0): void {
        $safe = $this->sanitizeName($name);
        $stream = @fopen($tmpPath, 'rb');
        if ($stream === false) {
            throw new \RuntimeException('Could not open downloaded temp file for "' . $name . '".');
        }
        try {
            try {
                $existing = $parent->get($safe);
                if ($existing instanceof Folder) {
                    throw new \RuntimeException('A folder named "' . $safe . '" already exists where a file was expected.');
                }
                $existing->putContent($stream);
                $node = $existing;
            } catch (NotFoundException) {
                $node = $parent->newFile($safe, $stream);
            }
            if ($mtime > 0 && $node instanceof Node) {
                try {
                    $node->touch($mtime);
                } catch (\Throwable) {
                    // Not critical if the storage does not support setting mtime.
                }
            }
        } finally {
            if (is_resource($stream)) {
                fclose($stream);
            }
        }
    }

    /**
     * Returns [size, mtime] for an existing child file, or null when it does
     * not exist (or is a folder).
     *
     * @return array{size: int, mtime: int}|null
     */
    public function getExistingFileInfo(Folder $parent, string $name): ?array {
        try {
            $node = $parent->get($this->sanitizeName($name));
        } catch (NotFoundException) {
            return null;
        }
        if ($node instanceof Folder) {
            return null;
        }
        return ['size' => (int)$node->getSize(), 'mtime' => (int)$node->getMTime()];
    }

    /**
     * @return list<string>
     */
    private function splitPath(string $relPath): array {
        return array_values(array_filter(
            array_map('trim', explode('/', str_replace('\\', '/', $relPath))),
            static fn (string $s): bool => $s !== '' && $s !== '.' && $s !== '..'
        ));
    }

    /** Replaces characters that are invalid on Nextcloud storages. */
    private function sanitizeName(string $name): string {
        $name = trim(str_replace(['/', '\\'], '-', $name));
        $name = preg_replace('/[\x00-\x1F]/', '', $name) ?? $name;
        return $name === '' ? '_' : $name;
    }
}
