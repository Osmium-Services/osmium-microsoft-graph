<?php

declare(strict_types=1);

namespace Osmium\Services\MicrosoftGraph\Models;

use Osmium\Core\Library\MailerInterface;

/**
 * Sends mail via the Microsoft Graph API using an Azure AD app
 * registration's client-credentials OAuth2 flow.
 *
 * Replaces SMTP for Microsoft 365 tenants that are disabling SMTP AUTH
 * org-wide - see https://aka.ms/smtp_auth_disabled. The app registration
 * needs the Mail.Send application permission with admin consent granted, and
 * should be scoped to the sending mailbox with an Exchange
 * ApplicationAccessPolicy so the client secret can only send as that one
 * address.
 *
 * Built by MailerFactory with the site config, which supplies the shared
 * From address and name; the tenant, client ID and secret come from this
 * service's own config. Plain curl, so there is no OAuth2 library to install.
 */
class MicrosoftGraphMailer implements MailerInterface
{
    private const TOKEN_URL_TEMPLATE = 'https://login.microsoftonline.com/%s/oauth2/v2.0/token';
    private const SEND_MAIL_URL_TEMPLATE = 'https://graph.microsoft.com/v1.0/users/%s/sendMail';
    private const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

    public function __construct(private object $config) {}

    /**
     * @param array<int, array{email: string, name?: string}> $recipients
     * @return array{success: bool, error?: string}
     */
    public function send(array $recipients, string $subject, string $htmlBody): array
    {
        try {
            $notReady = !MicrosoftGraphConfig::isReady();
            if ($notReady) throw new \RuntimeException('Microsoft 365 Mail is not configured: set the Tenant ID, Client ID and client secret on its settings page.');

            $token = $this->getAccessToken();
            $this->postSendMail($token, $recipients, $subject, $htmlBody);
            return ['success' => true];
        } catch (\Exception $e) {
            return ['success' => false, 'error' => $e->getMessage() . $this->hintForError($e->getMessage())];
        }
    }

    private function getAccessToken(): string
    {
        $graph = MicrosoftGraphConfig::get();
        $tokenUrl = \sprintf(self::TOKEN_URL_TEMPLATE, $graph->tenantId);

        $postFields = \http_build_query([
            'client_id' => $graph->clientId,
            'client_secret' => $graph->clientSecret,
            'scope' => self::GRAPH_SCOPE,
            'grant_type' => 'client_credentials',
        ]);

        $ch = \curl_init($tokenUrl);
        \curl_setopt_array($ch, [
            CURLOPT_POST => true,
            CURLOPT_POSTFIELDS => $postFields,
            CURLOPT_HTTPHEADER => ['Content-Type: application/x-www-form-urlencoded'],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT => 30,
        ]);
        $response = \curl_exec($ch);
        $curlError = \curl_error($ch);
        $httpCode = \curl_getinfo($ch, CURLINFO_HTTP_CODE);

        if ($curlError) throw new \RuntimeException("Graph token request curl error: {$curlError}");

        $decoded = \json_decode(json: (string) $response, associative: true);
        $tokenMissing = !isset($decoded['access_token']);
        if ($tokenMissing) {
            $error = $decoded['error_description'] ?? "HTTP {$httpCode}, empty response - check Tenant ID is set";
            throw new \RuntimeException("Graph token request failed: {$error}");
        }

        return $decoded['access_token'];
    }

    /**
     * @param array<int, array{email: string, name?: string}> $recipients
     */
    private function postSendMail(string $token, array $recipients, string $subject, string $htmlBody): void
    {
        $fromAddress = $this->config->email->fromAddress;
        $sendMailUrl = \sprintf(self::SEND_MAIL_URL_TEMPLATE, $fromAddress);

        $toRecipients = \array_map(
            fn (array $recipient) => ['emailAddress' => [
                'address' => $recipient['email'],
                'name' => $recipient['name'] ?? $recipient['email'],
            ]],
            $recipients,
        );

        $payload = [
            'message' => [
                'subject' => $subject,
                'body' => ['contentType' => 'HTML', 'content' => $htmlBody],
                'toRecipients' => $toRecipients,
                'from' => ['emailAddress' => [
                    'address' => $fromAddress,
                    'name' => $this->config->email->fromName ?? '',
                ]],
            ],
            'saveToSentItems' => false,
        ];

        $ch = \curl_init($sendMailUrl);
        \curl_setopt_array($ch, [
            CURLOPT_POST => true,
            CURLOPT_POSTFIELDS => \json_encode($payload),
            CURLOPT_HTTPHEADER => [
                'Authorization: Bearer ' . $token,
                'Content-Type: application/json',
            ],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT => 30,
        ]);
        $response = \curl_exec($ch);
        $curlError = \curl_error($ch);
        $httpCode = \curl_getinfo($ch, CURLINFO_HTTP_CODE);

        if ($curlError) throw new \RuntimeException("Graph sendMail curl error: {$curlError}");

        $failed = $httpCode < 200 || $httpCode >= 300;
        if ($failed) {
            $decoded = \json_decode(json: (string) $response, associative: true);
            $error = $decoded['error']['message'] ?? $response;
            throw new \RuntimeException("Graph sendMail failed ({$httpCode}): {$error}");
        }
    }

    /**
     * Turns the commonest Azure AD failures into a plain-English next step,
     * so the admin doesn't have to go AADSTS-code-spelunking.
     */
    private function hintForError(string $error): string
    {
        $isInvalidClientSecret = \str_contains($error, 'AADSTS7000215') || \str_contains($error, 'AADSTS7000222');
        if ($isInvalidClientSecret) {
            return "\n\nHint: the client secret saved here was rejected - it's either wrong or has "
                . 'expired. Generate a new one in the Azure AD app registration under Certificates & '
                . 'secrets, and save it here.';
        }

        $isInvalidClientOrTenant = \str_contains($error, 'AADSTS700016') || \str_contains($error, 'AADSTS90002');
        if ($isInvalidClientOrTenant) {
            return "\n\nHint: the Client ID or Tenant ID saved here doesn't match an app registration "
                . "Azure AD recognises. Double check both against the Azure AD app registration's "
                . 'Overview page.';
        }

        $isMissingConsent = \str_contains($error, 'AADSTS65001') || \str_contains($error, 'AADSTS500011');
        if ($isMissingConsent) {
            return "\n\nHint: the app registration hasn't been granted admin consent for the Mail.Send "
                . 'application permission. In Azure AD > App registrations > API permissions, add '
                . 'Microsoft Graph > Application permissions > Mail.Send, then click "Grant admin consent".';
        }

        $isAccessDenied = \str_contains($error, 'ErrorAccessDenied') || \str_contains($error, 'Forbidden');
        if ($isAccessDenied) {
            return "\n\nHint: the app is authenticated but isn't allowed to send as this mailbox. Check "
                . 'the Mail.Send permission has admin consent, and that no Exchange ApplicationAccessPolicy '
                . 'is scoping this app to a different mailbox.';
        }

        return '';
    }
}
