<?php

declare(strict_types=1);

/** @var array $_
 * @var bool $_['enabled']
 * @var string $_['tenantId']
 * @var string $_['clientId']
 * @var bool $_['hasClientSecret']
 * @var string $_['sourceUrl']
 * @var int $_['destGroupFolderId']
 * @var string $_['destSubPath']
 * @var string $_['runAsUser']
 * @var bool $_['replicateFullPath']
 * @var bool $_['connected']
 * @var string $_['accountName']
 * @var array $_['groupFolders']
 * @var string $_['redirectUri']
 * @var string $_['oauthStartUrl']
 */
script('neura_microsoft_sharepoint_sync', 'admin');
style('neura_microsoft_sharepoint_sync', 'admin');

function sps_checked(bool $val): string { return $val ? ' checked' : ''; }
?>

<div id="sps-root" class="section">

    <h2 class="sps-title">
        <svg class="sps-title-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7l9-4 9 4-9 4-9-4z"/><path d="M3 7v10l9 4 9-4V7"/><path d="M12 11v10"/></svg>
        Microsoft SharePoint File Sync
    </h2>
    <p class="sps-subtitle">Recursively download a SharePoint / OneDrive folder into a Nextcloud group folder, preserving the folder tree.</p>

    <!-- Setup guide -->
    <details class="sps-guide">
        <summary class="sps-guide__summary">
            <svg class="sps-guide__arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg>
            Setup guide — register an Azure AD app first
        </summary>
        <ol class="sps-guide__steps">
            <li><strong>Register an app</strong> in <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">Microsoft Entra ID → App registrations → New registration ↗</a>. Choose account type <em>Accounts in this organizational directory</em> (or multi-tenant).</li>
            <li><strong>Add a Web redirect URI</strong> pointing exactly to:<br><code class="sps-redirect"><?php p($_['redirectUri']); ?></code></li>
            <li><strong>API permissions → Add → Microsoft Graph → Delegated</strong>: add <code>Files.Read.All</code>, <code>Sites.Read.All</code>, <code>User.Read</code>, <code>offline_access</code>. Then <em>Grant admin consent</em>.</li>
            <li><strong>Certificates &amp; secrets → New client secret</strong>. Copy the secret <em>value</em>.</li>
            <li>Paste the <strong>Directory (tenant) ID</strong>, <strong>Application (client) ID</strong> and the <strong>client secret</strong> below, click <em>Save</em>, then <em>Connect Microsoft</em>.</li>
        </ol>
    </details>

    <div class="sps-card">

        <!-- Enable -->
        <div class="sps-field sps-field--inline">
            <label class="sps-toggle" for="sps-enabled">
                <input type="checkbox" id="sps-enabled"<?= sps_checked($_['enabled']) ?>>
                <span class="sps-toggle__track"></span>
                <span>Enable sync</span>
            </label>
        </div>

        <div class="sps-divider"></div>

        <!-- Azure app -->
        <h3 class="sps-section-title">Azure AD application</h3>

        <div class="sps-field">
            <label class="sps-label" for="sps-tenant">Directory (tenant) ID</label>
            <input class="sps-input" type="text" id="sps-tenant" value="<?php p($_['tenantId']); ?>" placeholder="common or a tenant GUID">
        </div>

        <div class="sps-field">
            <label class="sps-label" for="sps-client-id">Application (client) ID</label>
            <input class="sps-input" type="text" id="sps-client-id" value="<?php p($_['clientId']); ?>" placeholder="00000000-0000-0000-0000-000000000000">
        </div>

        <div class="sps-field">
            <label class="sps-label" for="sps-client-secret">Client secret</label>
            <div id="sps-secret-badge" class="sps-badge <?= $_['hasClientSecret'] ? 'sps-badge--ok' : 'sps-badge--warn' ?>">
                <?= $_['hasClientSecret'] ? '✓ Secret stored' : '⚠ No secret stored' ?>
            </div>
            <input class="sps-input" type="password" id="sps-client-secret" value="" placeholder="<?= $_['hasClientSecret'] ? 'Leave empty to keep current secret' : 'Paste the client secret value' ?>" autocomplete="new-password">
        </div>

        <div class="sps-divider"></div>

        <!-- Connection -->
        <h3 class="sps-section-title">Microsoft connection</h3>
        <div class="sps-field">
            <div id="sps-conn-badge" class="sps-badge <?= $_['connected'] ? 'sps-badge--ok' : 'sps-badge--warn' ?>">
                <?= $_['connected'] ? '✓ Connected' . ($_['accountName'] !== '' ? ' — ' . \htmlspecialchars($_['accountName'], ENT_QUOTES) : '') : '⚠ Not connected' ?>
            </div>
            <div class="sps-actions">
                <a class="sps-btn sps-btn--primary" href="<?php p($_['oauthStartUrl']); ?>">Connect Microsoft</a>
                <button type="button" id="sps-disconnect" class="sps-btn sps-btn--ghost"<?= $_['connected'] ? '' : ' style="display:none"' ?>>Disconnect</button>
            </div>
            <span class="sps-hint">Save the Azure app fields before connecting. You will be redirected to Microsoft to sign in and consent.</span>
        </div>

        <div class="sps-divider"></div>

        <!-- Source -->
        <h3 class="sps-section-title">Source</h3>
        <div class="sps-field">
            <label class="sps-label" for="sps-source-url">SharePoint folder URL</label>
            <input class="sps-input" type="text" id="sps-source-url" value="<?php p($_['sourceUrl']); ?>" placeholder="https://contoso.sharepoint.com/sites/SITE/LIBRARY/FOLDER/SUBFOLDER">
            <span class="sps-hint">Paste the browser URL of the shared folder. All files and sub-folders below it are downloaded.</span>
        </div>

        <div class="sps-divider"></div>

        <!-- Destination -->
        <h3 class="sps-section-title">Destination</h3>
        <div class="sps-field">
            <label class="sps-label" for="sps-groupfolder">Group folder</label>
            <select class="sps-input" id="sps-groupfolder">
                <?php if (empty($_['groupFolders'])): ?>
                    <option value="0">No group folders found — install the "Group folders" app</option>
                <?php else: ?>
                    <option value="0">— Select a group folder —</option>
                    <?php foreach ($_['groupFolders'] as $gf): ?>
                        <option value="<?php p((string)$gf['id']); ?>"<?= (int)$gf['id'] === (int)$_['destGroupFolderId'] ? ' selected' : '' ?>><?php p($gf['mountPoint']); ?></option>
                    <?php endforeach; ?>
                <?php endif; ?>
            </select>
        </div>

        <div class="sps-field">
            <label class="sps-label" for="sps-subpath">Sub-path inside the group folder <span class="sps-optional">(optional)</span></label>
            <input class="sps-input" type="text" id="sps-subpath" value="<?php p($_['destSubPath']); ?>" placeholder="e.g. Imports/SharePoint">
        </div>

        <div class="sps-field">
            <label class="sps-label" for="sps-runas">Write as user</label>
            <input class="sps-input" type="text" id="sps-runas" value="<?php p($_['runAsUser']); ?>" placeholder="a user with write access to the group folder">
            <span class="sps-hint">This Nextcloud user's access is used to write into the group folder. They must belong to a group with write permission on it.</span>
        </div>

        <div class="sps-field sps-field--inline">
            <label class="sps-toggle" for="sps-replicate">
                <input type="checkbox" id="sps-replicate"<?= sps_checked($_['replicateFullPath']) ?>>
                <span class="sps-toggle__track"></span>
                <span>Replicate the full SharePoint path</span>
            </label>
        </div>
        <span class="sps-hint">On: recreate the whole folder path (e.g. <code>Campaigns/2026/Q1-Report</code>). Off: create only the selected leaf folder.</span>

        <div class="sps-divider"></div>

        <!-- Actions -->
        <div class="sps-actions">
            <button id="sps-save" class="sps-btn sps-btn--primary" type="button">Save settings</button>
            <button id="sps-test" class="sps-btn sps-btn--secondary" type="button">Test connection</button>
            <button id="sps-sync-now" class="sps-btn sps-btn--secondary" type="button">Download now</button>
        </div>

        <div id="sps-status" class="sps-status-msg" style="display:none"></div>

    </div><!-- .sps-card -->

</div><!-- #sps-root -->
