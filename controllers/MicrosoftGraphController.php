<?php

declare(strict_types=1);

namespace Osmium\Services\MicrosoftGraph\Controllers;

use Osmium\Modules\Admin\Core\AdminController;
use Osmium\Services\MicrosoftGraph\Models\MicrosoftGraphConfig;

/**
 * Microsoft 365 Mail settings controller - full-page form POST/redirect, same
 * shape as TurnstileController.
 *
 * Routes:
 *   - index() → /admin/settings/microsoft-graph/  (GET shows the form, POST saves it)
 */
class MicrosoftGraphController extends AdminController
{
    private const CONFIG_FILE_PATH = 'app/config/services/microsoft-graph.json.php';

    public function index(): void
    {
        $isPost = $this->isPost();
        if ($isPost) $this->handleSubmit();

        $this->data['admin']['config']['microsoftGraph'] = (array) MicrosoftGraphConfig::get();
        $this->data['admin']['settingsSaved'] = $_SESSION['microsoft_graph_settings_saved'] ?? false;
        $this->data['admin']['settingsError'] = $_SESSION['microsoft_graph_settings_error'] ?? false;
        unset($_SESSION['microsoft_graph_settings_saved'], $_SESSION['microsoft_graph_settings_error']);

        $this->setView('microsoft-graph/index.phtml');
    }

    private function handleSubmit(): void
    {
        $csrfValid = $this->admin->auth->validateCsrf();
        if (!$csrfValid) {
            $_SESSION['microsoft_graph_settings_error'] = 'Invalid form submission. Please try again.';
            $this->redirect('settings/microsoft-graph/');
        }

        $tenantId = \trim($_POST['tenant_id'] ?? '');
        $clientId = \trim($_POST['client_id'] ?? '');

        $postedSecret = \trim($_POST['client_secret'] ?? '');
        $clientSecret = $postedSecret === '' ? (string) (MicrosoftGraphConfig::get()->clientSecret ?? '') : $postedSecret; // Blank keeps the stored secret

        $this->saveConfig($tenantId, $clientId, $clientSecret);

        $this->admin->model->changelog->log(
            description: 'Updated Microsoft 365 Mail settings',
            recordType: 'settings',
        );

        MicrosoftGraphConfig::clearCache();

        $_SESSION['microsoft_graph_settings_saved'] = true;
        $this->redirect('settings/microsoft-graph/');
    }

    private function saveConfig(string $tenantId, string $clientId, string $clientSecret): void
    {
        $configExists = \file_exists(self::CONFIG_FILE_PATH);
        if (!$configExists) $this->ensureConfigDirectoryExists();

        $newJson = \json_encode(
            value: ['microsoftGraph' => [
                'tenantId' => $tenantId,
                'clientId' => $clientId,
                'clientSecret' => $clientSecret,
            ]],
            flags: JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE,
        );

        \file_put_contents(self::CONFIG_FILE_PATH, "<?php exit(); ?>\n" . $newJson . "\n");
    }

    private function ensureConfigDirectoryExists(): void
    {
        $dir = \dirname(self::CONFIG_FILE_PATH);
        $alreadyExists = \is_dir($dir);
        if (!$alreadyExists) \mkdir(directory: $dir, permissions: 0755, recursive: true);
    }

    private function isPost(): bool
    {
        return $_SERVER['REQUEST_METHOD'] === 'POST';
    }
}
