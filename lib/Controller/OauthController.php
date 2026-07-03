<?php

declare(strict_types=1);

namespace OCA\NeuraMicrosoftSharepointSync\Controller;

use OCA\NeuraMicrosoftSharepointSync\AppInfo\Application;
use OCA\NeuraMicrosoftSharepointSync\Service\ConfigService;
use OCA\NeuraMicrosoftSharepointSync\Service\MsGraphService;
use OCP\AppFramework\Controller;
use OCP\AppFramework\Http\RedirectResponse;
use OCP\IRequest;
use OCP\ISession;
use OCP\IURLGenerator;
use OCP\Security\ISecureRandom;
use Psr\Log\LoggerInterface;

/**
 * Handles the browser-facing part of the delegated OAuth2 flow:
 * redirecting the admin to Microsoft and receiving the authorization code back.
 *
 * CSRF is protected via the OAuth "state" parameter stored in the session,
 * because the callback is an external redirect without a Nextcloud request token.
 */
class OauthController extends Controller {

    private const SESSION_STATE_KEY = 'neura_sp_oauth_state';

    public function __construct(
        IRequest $request,
        private ConfigService $configService,
        private MsGraphService $graphService,
        private ISession $session,
        private IURLGenerator $urlGenerator,
        private ISecureRandom $secureRandom,
        private LoggerInterface $logger,
    ) {
        parent::__construct(Application::APP_ID, $request);
    }

    private function callbackUrl(): string {
        return $this->urlGenerator->linkToRouteAbsolute(Application::APP_ID . '.oauth.callback');
    }

    private function settingsUrl(string $fragment): string {
        return $this->urlGenerator->linkToRoute('settings.AdminSettings.index', ['section' => Application::APP_ID]) . '#' . $fragment;
    }

    /**
     * Starts the OAuth flow by redirecting the admin to the Microsoft sign-in page.
     *
     * @AdminRequired
     * @NoCSRFRequired
     * @UseSession
     */
    public function start(): RedirectResponse {
        if (!$this->configService->hasAppCredentials()) {
            return new RedirectResponse($this->settingsUrl('error-nocreds'));
        }
        $state = $this->secureRandom->generate(32, ISecureRandom::CHAR_ALPHANUMERIC);
        $this->session->set(self::SESSION_STATE_KEY, $state);
        return new RedirectResponse($this->graphService->getAuthorizeUrl($state, $this->callbackUrl()));
    }

    /**
     * OAuth redirect target. Verifies the state, exchanges the code for tokens,
     * and returns the admin to the settings page.
     *
     * @AdminRequired
     * @NoCSRFRequired
     * @UseSession
     */
    public function callback(string $code = '', string $state = '', string $error = '', string $error_description = ''): RedirectResponse {
        $expectedState = (string)$this->session->get(self::SESSION_STATE_KEY);
        $this->session->remove(self::SESSION_STATE_KEY);

        if ($error !== '') {
            $this->logger->warning('OAuth error: ' . $error . ' ' . $error_description, ['app' => Application::APP_ID]);
            return new RedirectResponse($this->settingsUrl('error-oauth'));
        }
        if ($state === '' || !hash_equals($expectedState, $state)) {
            return new RedirectResponse($this->settingsUrl('error-state'));
        }
        if ($code === '') {
            return new RedirectResponse($this->settingsUrl('error-nocode'));
        }

        try {
            $this->graphService->exchangeCode($code, $this->callbackUrl());
        } catch (\Throwable $e) {
            $this->logger->error('Token exchange failed: ' . $e->getMessage(), ['app' => Application::APP_ID]);
            return new RedirectResponse($this->settingsUrl('error-token'));
        }

        return new RedirectResponse($this->settingsUrl('connected'));
    }
}
