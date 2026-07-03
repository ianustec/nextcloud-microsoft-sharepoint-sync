<?php

declare(strict_types=1);

return [
    'routes' => [
        [
            'name' => 'AdminSettings#save',
            'url' => '/api/admin/settings',
            'verb' => 'POST',
        ],
        [
            'name' => 'AdminSettings#testConnection',
            'url' => '/api/admin/test',
            'verb' => 'POST',
        ],
        [
            'name' => 'AdminSettings#listGroupFolders',
            'url' => '/api/admin/groupfolders',
            'verb' => 'GET',
        ],
        [
            'name' => 'AdminSettings#syncNow',
            'url' => '/api/admin/sync-now',
            'verb' => 'POST',
        ],
        [
            'name' => 'AdminSettings#disconnect',
            'url' => '/api/admin/disconnect',
            'verb' => 'POST',
        ],
        [
            'name' => 'oauth#start',
            'url' => '/oauth/start',
            'verb' => 'GET',
        ],
        [
            'name' => 'oauth#callback',
            'url' => '/oauth/callback',
            'verb' => 'GET',
        ],
    ],
];
