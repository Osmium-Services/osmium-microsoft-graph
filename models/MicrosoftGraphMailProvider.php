<?php

declare(strict_types=1);

namespace Osmium\Services\MicrosoftGraph\Models;

/**
 * Adapts Microsoft 365 mail to core's mail.providers hook (see ServiceHooks /
 * MailerFactory).
 */
class MicrosoftGraphMailProvider
{
    public const ID = 'microsoft-graph';

    /**
     * mail.providers - advertise Microsoft 365 Mail and whether it can send now.
     */
    public static function provider(array $payload): array
    {
        return [
            'id' => self::ID,
            'label' => 'Microsoft 365 (Graph API)',
            'ready' => MicrosoftGraphConfig::isReady(),
            'settingsRoute' => 'settings/microsoft-graph/',
            'mailer' => MicrosoftGraphMailer::class,
        ];
    }
}
