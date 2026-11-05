<?php
declare(strict_types=1);
require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/database.php';
require __DIR__ . '/inventory.php';

$error = '';
$message = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        $file = $_FILES['stock'] ?? null;
        if (!$file || ($file['error'] ?? null) !== UPLOAD_ERR_OK
            || !is_string($file['tmp_name'] ?? null) || !is_uploaded_file($file['tmp_name'])) {
            throw new RuntimeException('Choose a stock Excel workbook to upload.');
        }
        if (filesize($file['tmp_name']) > 10 * 1024 * 1024) {
            throw new RuntimeException('Maximum workbook size is 10 MB.');
        }
        $products = stockProductsFromWorkbook($file['tmp_name']);
        $count = stockImport(stockDb(), $products);
        $message = $count . ' products imported. Existing products were updated by their main SKU.';
    } catch (Throwable $exception) {
        $error = $exception->getMessage();
    }
}
function h(string $value): string { return htmlspecialchars($value, ENT_QUOTES, 'UTF-8'); }
?>
<!doctype html>
<html lang="en">
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Import product stock</title>
<body style="font-family:Arial,sans-serif;max-width:720px;margin:40px auto;padding:16px">
<h1>Import product stock</h1>
<p>Upload StockCount_FinishedGoods.xlsx using its original column layout.</p>
<?php if ($error !== ''): ?><p role="alert" style="color:#991b1b"><?= h($error) ?></p><?php endif; ?>
<?php if ($message !== ''): ?><p role="status"><?= h($message) ?></p><?php endif; ?>
<form method="post" enctype="multipart/form-data">
<label>Stock workbook <input type="file" name="stock" accept=".xlsx,.xls" required></label>
<button type="submit">Import products</button>
</form>
<p><a href="bundle_rules.php">Manage bundle rules</a> | <a href="converter.php">Convert courier Excel</a></p>
</body>
</html>

