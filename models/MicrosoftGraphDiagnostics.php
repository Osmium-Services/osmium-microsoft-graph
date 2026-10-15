<?php

declare(strict_types=1);

namespace Osmium\Services\MicrosoftGraph\Models;

use Osmium\Core\Library\MailDiagnosticsInterface;

/**
 * What the Email settings "Run diagnostics" button checks for Microsoft 365
 * Mail, in place of the SMTP connection checks: the credentials are saved,
 * Azure AD hands out a token for them, and that token carries the Mail.Send
 * application permission.
 *
 * It cannot see whether the mailbox exists or an Exchange
 * ApplicationAccessPolicy blocks it (that needs extra Graph permissions), so
 * "Send test email" remains the proof for those.
 */
class MicrosoftGraphDiagnostics implements MailDiagnosticsInterface
{
    public function __construct(private object $config) {}

    /**
     * @return array<int, array{id: string, label: string, status: string, detail: string, fix: string}>
     */
    public function checks(): array
    {
        $notReady = !MicrosoftGraphConfig::isReady();
        if ($notReady) {
            return [$this->row('graph-config', 'Microsoft 365 credentials', 'fail', 'The Tenant ID, Client ID or client secret is not set.', 'Enter them on the Microsoft 365 Mail settings page.')];
        }

        $mailer = new MicrosoftGraphMailer($this->config);

        try {
            $token = $mailer->getAccessToken();
        } catch (\Exception $e) {
            $message = $e->getMessage();
            $hint = \trim($mailer->hintForError($message));
            return [
                $this->row('graph-config', 'Microsoft 365 credentials', 'pass', 'Tenant ID, Client ID and client secret are set.'),
                $this->row('graph-token', 'Microsoft 365 sign-in', 'fail', $message, $hint),
            ];
        }

        return [
            $this->row('graph-config', 'Microsoft 365 credentials', 'pass', 'Tenant ID, Client ID and client secret are set.'),
            $this->row('graph-token', 'Microsoft 365 sign-in', 'pass', 'Azure AD accepted the credentials and issued a token.'),
            $this->permissionRow($token),
            $this->row('graph-mailbox', 'Sending mailbox', 'skip', 'Whether the From mailbox exists, and is not blocked by an ApplicationAccessPolicy, cannot be checked without extra permissions.', 'Use "Send test email" to prove it.'),
        ];
    }

    /**
     * Reads the roles claim of the (unverified, just received) access token to
     * see whether admin consent for Mail.Send was granted.
     *
     * @return array{id: string, label: string, status: string, detail: string, fix: string}
     */
    private function permissionRow(string $token): array
    {
        $parts = \explode('.', $token);
        $payload = $parts[1] ?? '';
        $json = \base64_decode(\strtr($payload, '-_', '+/'), true);
        $claims = $json === false ? null : \json_decode($json, true);

        $unreadable = !\is_array($claims);
        if ($unreadable) {
            return $this->row('graph-permission', 'Mail.Send permission', 'skip', 'The token could not be read, so the permission was not checked.');
        }

        $roles = (array)($claims['roles'] ?? []);
        $granted = \in_array('Mail.Send', $roles, true);
        if ($granted) {
            return $this->row('graph-permission', 'Mail.Send permission', 'pass', 'The app has the Mail.Send application permission.');
        }

        return $this->row(
            'graph-permission',
            'Mail.Send permission',
            'fail',
            'The token does not carry the Mail.Send application permission.',
            'In Azure AD > App registrations > API permissions, add Microsoft Graph > Application permissions > Mail.Send, then click "Grant admin consent".',
        );
    }

    /** @return array{id: string, label: string, status: string, detail: string, fix: string} */
    private function row(string $id, string $label, string $status, string $detail, string $fix = ''): array
    {
        return ['id' => $id, 'label' => $label, 'status' => $status, 'detail' => $detail, 'fix' => $fix];
    }
}
