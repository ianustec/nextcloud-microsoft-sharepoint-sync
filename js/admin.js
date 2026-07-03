(function () {
    const APP_URL = OC.generateUrl('/apps/neura_microsoft_sharepoint_sync');

    const saveBtn       = document.getElementById('sps-save');
    const testBtn       = document.getElementById('sps-test');
    const syncBtn       = document.getElementById('sps-sync-now');
    const disconnectBtn = document.getElementById('sps-disconnect');
    const statusDiv     = document.getElementById('sps-status');

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = 'sps-status-msg sps-status-msg--' + (type || 'ok');
        statusDiv.style.display = 'block';
    }

    function setBtnBusy(btn, label) {
        btn.disabled = true;
        btn._origText = btn.textContent;
        btn.textContent = label;
    }

    function resetBtn(btn) {
        btn.disabled = false;
        btn.textContent = btn._origText;
    }

    function escHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    // Show a message based on the OAuth callback redirect fragment.
    (function handleOauthResult() {
        const hash = window.location.hash;
        if (!hash) return;
        const map = {
            '#connected': ['Microsoft account connected successfully.', 'ok'],
            '#error-nocreds': ['Save the Azure app fields (tenant, client ID, secret) before connecting.', 'error'],
            '#error-oauth': ['Microsoft denied the authorization request.', 'error'],
            '#error-state': ['Security check failed (state mismatch). Please try connecting again.', 'error'],
            '#error-nocode': ['No authorization code was returned by Microsoft.', 'error'],
            '#error-token': ['Could not exchange the code for a token. Check the client secret and redirect URI.', 'error'],
        };
        if (map[hash]) {
            showStatus(map[hash][0], map[hash][1]);
            history.replaceState(null, '', window.location.pathname + window.location.search);
        }
    })();

    function collectPayload() {
        return {
            enabled: document.getElementById('sps-enabled').checked ? 'yes' : 'no',
            tenantId: document.getElementById('sps-tenant').value.trim() || 'common',
            clientId: document.getElementById('sps-client-id').value.trim(),
            clientSecret: document.getElementById('sps-client-secret').value,
            sourceUrl: document.getElementById('sps-source-url').value.trim(),
            destGroupFolderId: parseInt(document.getElementById('sps-groupfolder').value, 10) || 0,
            destSubPath: document.getElementById('sps-subpath').value.trim(),
            runAsUser: document.getElementById('sps-runas').value.trim(),
            replicateFullPath: document.getElementById('sps-replicate').checked ? 'yes' : 'no',
        };
    }

    saveBtn.addEventListener('click', async function () {
        setBtnBusy(saveBtn, 'Saving…');
        statusDiv.style.display = 'none';
        try {
            const resp = await fetch(APP_URL + '/api/admin/settings', {
                method: 'POST',
                headers: { requesttoken: OC.requestToken, 'Content-Type': 'application/json' },
                body: JSON.stringify(collectPayload()),
            });
            const json = await resp.json();
            if (resp.ok) {
                showStatus('Settings saved.', 'ok');
                document.getElementById('sps-client-secret').value = '';
            } else {
                showStatus(json.message || 'Save failed.', 'error');
            }
        } catch (err) {
            showStatus(err.message, 'error');
        } finally {
            resetBtn(saveBtn);
        }
    });

    testBtn.addEventListener('click', async function () {
        setBtnBusy(testBtn, 'Testing…');
        statusDiv.style.display = 'none';
        try {
            const resp = await fetch(APP_URL + '/api/admin/test', {
                method: 'POST',
                headers: { requesttoken: OC.requestToken, 'Content-Type': 'application/json' },
                body: '{}',
            });
            const json = await resp.json();
            if (resp.ok && json.status === 'ok') {
                showStatus('Connection OK — ' + (json.account || 'account reachable'), 'ok');
            } else {
                showStatus(json.message || 'Connection failed.', 'error');
            }
        } catch (err) {
            showStatus(err.message, 'error');
        } finally {
            resetBtn(testBtn);
        }
    });

    if (disconnectBtn) {
        disconnectBtn.addEventListener('click', async function () {
            if (!confirm('Disconnect the Microsoft account? You will need to reconnect to sync again.')) return;
            setBtnBusy(disconnectBtn, 'Disconnecting…');
            try {
                await fetch(APP_URL + '/api/admin/disconnect', {
                    method: 'POST',
                    headers: { requesttoken: OC.requestToken },
                });
                window.location.reload();
            } catch (err) {
                showStatus(err.message, 'error');
                resetBtn(disconnectBtn);
            }
        });
    }

    syncBtn.addEventListener('click', async function () {
        if (!confirm('Start downloading from SharePoint now? Large folders can take a while.')) return;
        setBtnBusy(syncBtn, 'Downloading…');
        statusDiv.className = 'sps-status-msg sps-status-msg--info';
        statusDiv.style.display = 'block';
        statusDiv.innerHTML = '<span class="sps-spinner"></span> Downloading… keep this tab open.';
        const runPayload = {
            sourceUrl: document.getElementById('sps-source-url').value.trim(),
            destGroupFolderId: parseInt(document.getElementById('sps-groupfolder').value, 10) || 0,
            destSubPath: document.getElementById('sps-subpath').value.trim(),
            runAsUser: document.getElementById('sps-runas').value.trim(),
            replicateFullPath: document.getElementById('sps-replicate').checked ? 'yes' : 'no',
        };
        try {
            const resp = await fetch(APP_URL + '/api/admin/sync-now', {
                method: 'POST',
                headers: { requesttoken: OC.requestToken, 'Content-Type': 'application/json' },
                body: JSON.stringify(runPayload),
            });
            const json = await resp.json();
            if (json.status === 'ok') {
                renderReport(json);
            } else if (json.status === 'skipped') {
                showStatus('Skipped: ' + (json.message || ''), 'info');
            } else {
                showStatus('Error: ' + (json.message || 'Download failed.'), 'error');
            }
        } catch (err) {
            showStatus(err.message, 'error');
        } finally {
            resetBtn(syncBtn);
        }
    });

    function renderReport(json) {
        const hasErrors = (json.failed || 0) > 0;
        statusDiv.className = 'sps-result-wrap';
        statusDiv.style.display = 'block';

        let html = '<div class="sps-pills">'
            + pill('Downloaded', json.downloaded || 0, 'ok')
            + pill('Skipped', json.skipped || 0, 'neutral')
            + pill('Folders', json.folders || 0, 'ok')
            + pill('Failed', json.failed || 0, hasErrors ? 'error' : 'ok')
            + '</div>';

        if (json.root) {
            html += '<p class="sps-hint">Destination sub-tree: <code>' + escHtml(json.root) + '</code></p>';
        }

        if (hasErrors && json.errors && json.errors.length) {
            html += '<details class="sps-detail sps-detail--error" open><summary>✗ Errors (' + json.errors.length + ')</summary><ul class="sps-err-list">';
            json.errors.forEach(function (e) {
                html += '<li>' + escHtml(e) + '</li>';
            });
            html += '</ul></details>';
        }

        statusDiv.innerHTML = html;
    }

    function pill(label, value, type) {
        return '<div class="sps-pill sps-pill--' + type + '">'
            + '<span class="sps-pill__val">' + value + '</span>'
            + '<span class="sps-pill__label">' + label + '</span>'
            + '</div>';
    }
})();
