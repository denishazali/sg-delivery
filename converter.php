<?php
declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/database.php';
require __DIR__ . '/inventory.php';

/*
 * EDIT: DEFAULT VALUES
 * Set fixed values for every output row here, e.g. 'Pickup Name' => 'Your Company'.
 * Quantity, Item Description and Unit Price are calculated later; edit their final assignments below.
 *
 * Configure information that is unavailable in the source Excel.
 * Blank defaults produce blank output cells.
 * Do not use sample weights, dimensions, or pickup details unless accurate.
 */
$defaults = [
    'Service Type'          => '',
    'Receiver Province'     => '',
    'Pickup Name'           => '',
    'Pickup Contact Number' => '',
    'Pickup Postal Code'    => '',
    'Pickup Address'        => '',
    'Pickup Instructions'   => '',
    'Recipient mailbox'     => '',
    'Delivery Instructions' => '',
    'Length (cm)'           => '',
    'Width (cm)'            => '',
    'Height (cm)'           => '',
    'Weight (kg)'           => '',
    'Insure Amount'         => '',
    'Item Weight (kg)'      => '',
    'Reference Number'      => '',
    'Item Value'            => '',
    'Goods Name'            => '',
];

/* EDIT: EXCEL HEADERS AND COLUMN ORDER
 * These names are also used by $defaults, $mapping and $record. Update those keys too if renaming a header.
 */
$templateHeaders = [
    'Service Type',
    'Receiver Country',
    'Receiver Province',
    'Receiver City',
    'Receiver State',
    'Reference Number',
    'Pickup Name',
    'Pickup Contact Number',
    'Pickup Postal Code',
    'Pickup Address',
    'Pickup Instructions',
    'Receiver Name',
    'Receiver Contact Number',
    'Receiver Postal Code',
    'Receiver Address',
    'Recipient mailbox',
    'Delivery Instructions',
    'Length (cm)',
    'Width (cm)',
    'Height (cm)',
    'Weight (kg)',
    'Item Value',
    'Item Description',
    'Insure Amount',
    'Goods Name',
    'Unit Price',
    'Item Weight (kg)',
    'Quantity',
];

/* EDIT: SOURCE COLUMN MAPPING
 * Left = output template column; right = uploaded Excel column.
 * Item Description, Quantity and Unit Price start here but are replaced by calculated values below.
 */
$mapping = [
    'Receiver Country'        => 'RECEIVER COUNTRY',
    'Receiver City'           => 'RECEIVER CITY',
    'Receiver State'          => 'RECEIVER STATE',
    'Receiver Name'           => 'RECEIVER NAME',
    'Receiver Contact Number' => 'RECEIVER CONTACT NO.',
    'Receiver Postal Code'    => 'RECEIVER POSTCODE',
    'Item Description'        => 'SKU DESCRIPTION',
    'Unit Price'              => 'UNIT SALE PRICE',
    'Quantity'                => 'ORDER QTY',
    'Reference Number'        => 'DOCUMENT NO',
];

function normalizeHeader(string $value): string
{
    $value = str_replace("\xC2\xA0", ' ', $value);
    $value = preg_replace('/\s+/u', ' ', $value) ?? $value;

    return strtoupper(trim($value));
}

function escapeHtml(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES, 'UTF-8');
}

/*
 * Read formatted values to preserve identifiers such as postal codes.
 * Reject formulas instead of calculating them during conversion.
 */
function readSourceCell(
    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet,
    int $column,
    int $row,
    bool $raw = false
): string {
    $address = Coordinate::stringFromColumnIndex($column) . $row;
    $cell = $sheet->getCell($address);

    if ($cell->getDataType() === DataType::TYPE_FORMULA) {
        throw new RuntimeException(
            "Cell {$address} contains a formula. Paste its value before uploading."
        );
    }

    return trim((string) ($raw ? $cell->getValue() : $cell->getFormattedValue()));
}

function readNumber(string $value, string $label, int $row): ?float
{
    if ($value === '') {
        return null;
    }

    if (!is_numeric($value)) {
        throw new RuntimeException(
            "Row {$row}: {$label} must contain a number without currency symbols."
        );
    }

    $number = (float) $value;

    if (!is_finite($number) || $number < 0) {
        throw new RuntimeException(
            "Row {$row}: {$label} must be a valid non-negative number."
        );
    }

    return $number;
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $input = null;
    $output = null;
    $downloadPath = null;

    try {
        $file = $_FILES['excel'] ?? null;

        if (!$file || is_array($file['error'] ?? null)) {
            throw new RuntimeException('Please select an Excel file.');
        }

        if ($file['error'] !== UPLOAD_ERR_OK) {
            throw new RuntimeException(
                'Upload failed. Check the file size and your PHP upload limits.'
            );
        }

        if (!is_uploaded_file($file['tmp_name'])) {
            throw new RuntimeException('Invalid uploaded file.');
        }

        $fileSize = filesize($file['tmp_name']);

        if ($fileSize === false || $fileSize > 10 * 1024 * 1024) {
            throw new RuntimeException('The maximum upload size is 10 MB.');
        }

        $extension = strtolower(
            pathinfo((string) $file['name'], PATHINFO_EXTENSION)
        );

        if (!in_array($extension, ['xlsx', 'xls'], true)) {
            throw new RuntimeException('Only .xlsx and .xls files are supported.');
        }

        // Check the actual format, rather than trusting the filename.
        $type = IOFactory::identify($file['tmp_name'], ['Xlsx', 'Xls']);
        $reader = IOFactory::createReader($type);

        // Preserve number formats, including leading-zero postal codes.
        $reader->setReadDataOnly(false);

        $input = $reader->load($file['tmp_name']);
        $source = $input->getSheet(0);

        $lastRow = $source->getHighestDataRow();
        $lastColumn = Coordinate::columnIndexFromString(
            $source->getHighestDataColumn()
        );

        if ($lastRow > 20001 || $lastColumn > 200) {
            throw new RuntimeException(
                'Please upload a sheet with at most 20,000 data rows and 200 columns.'
            );
        }

        // The first worksheet must have its headers in row 1.
        $sourceColumns = [];

        for ($column = 1; $column <= $lastColumn; $column++) {
            $header = normalizeHeader(
                readSourceCell($source, $column, 1)
            );

            if ($header === '') {
                continue;
            }

            if (isset($sourceColumns[$header])) {
                throw new RuntimeException("Duplicate source header: {$header}");
            }

            $sourceColumns[$header] = $column;
        }

        $requiredHeaders = array_unique(array_merge(
            array_values($mapping),
            [
                'RECEIVER ADDRESS 1',
                'RECEIVER ADDRESS 2',
                'RECEIVER ADDRESS 3',
                'CLIENT SKU NO.',
            ]
        ));

        foreach ($requiredHeaders as $header) {
            if (!isset($sourceColumns[normalizeHeader($header)])) {
                throw new RuntimeException(
                    "Missing source column: {$header}. Headers must be in row 1."
                );
            }
        }
        $readHeaders = array_unique(array_merge(
            $requiredHeaders,
            array_intersect(['DOCUMENT NO', 'STORAGE CLIENT NO'], array_keys($sourceColumns))
        ));

        $catalog = stockLoadCatalog(stockDb());
        $output = new Spreadsheet();
        $target = $output->getActiveSheet();
        $target->setTitle('Courier Template');

        foreach ($templateHeaders as $index => $header) {
            $address = Coordinate::stringFromColumnIndex($index + 1) . '1';
            $target->setCellValueExplicit(
                $address,
                $header,
                DataType::TYPE_STRING
            );
        }

        $numericHeaders = [
            'Length (cm)',
            'Width (cm)',
            'Height (cm)',
            'Weight (kg)',
            'Item Value',
            'Insure Amount',
            'Unit Price',
            'Item Weight (kg)',
            'Quantity',
        ];

        $outputRow = 2;
        $reviewRows = [];
        $orders = [];

        for ($row = 2; $row <= $lastRow; $row++) {
            $values = [];

            foreach ($readHeaders as $header) {
                $header = normalizeHeader($header);
                $values[$header] = readSourceCell(
                    $source,
                    $sourceColumns[$header],
                    $row,
                    in_array($header, ['ORDER QTY', 'UNIT SALE PRICE'], true)
                );
            }

            // Skip completely empty source rows.
            if (!array_filter(
                $values,
                static fn(string $value): bool => $value !== ''
            )) {
                continue;
            }

            $get = static fn(string $header): string =>
                $values[normalizeHeader($header)] ?? '';

            $record = array_fill_keys($templateHeaders, '');
            $record = array_replace($record, $defaults);

            foreach ($mapping as $destination => $sourceHeader) {
                $record[$destination] = $get($sourceHeader);
            }

            // EDIT: RECEIVER ADDRESS — combine source address lines; change ', ' to change the separator.
            $record['Receiver Address'] = implode(', ', array_filter(
                [
                    $get('RECEIVER ADDRESS 1'),
                    $get('RECEIVER ADDRESS 2'),
                    $get('RECEIVER ADDRESS 3'),
                ],
                static fn(string $part): bool => $part !== ''
            ));

            $quantity = readNumber($get('ORDER QTY'), 'ORDER QTY', $row);
            $unitPrice = readNumber(
                $get('UNIT SALE PRICE'),
                'UNIT SALE PRICE',
                $row
            );

            if ($quantity === null || $quantity < 1
                || $quantity > 999999999 || floor($quantity) !== $quantity) {
                throw new RuntimeException("Row {$row}: ORDER QTY must be a positive whole number.");
            }
            $reviewReason = null;
            try {
                $match = stockResolve($catalog, $get('CLIENT SKU NO.'), $get('SKU DESCRIPTION'), $row);
            } catch (StockItemReviewException $exception) {
                // Preserve the original item when its product/carton split cannot be determined.
                $reviewReason = $exception->getMessage();
                $match = null;
            }

            /*
             * Restore a dropped leading zero for Singapore postal codes.
             * Phone numbers and reference numbers remain text.
             */
            $country = strtoupper(trim($record['Receiver Country']));
            $postcode = $record['Receiver Postal Code'];

            if (
                in_array($country, ['SINGAPORE', 'SG', 'SGP'], true)
                && preg_match('/^\d{1,6}$/D', $postcode)
            ) {
                $record['Receiver Postal Code'] = str_pad(
                    $postcode,
                    6,
                    '0',
                    STR_PAD_LEFT
                );
            }

            // EDIT: ORDER GROUPING — combine rows only when the document, client and delivery details agree.
            // Reference Number uses DOCUMENT NO via $mapping; missing document numbers keep rows separate.
            $document = $get('DOCUMENT NO');
            $orderIdentity = [$document !== '', $document === '' ? $row : stockKey($document),
                stockKey($get('STORAGE CLIENT NO'))];
            foreach (['Receiver Name', 'Receiver Contact Number', 'Receiver Postal Code',
                      'Receiver Address', 'Receiver Country', 'Receiver City', 'Receiver State'] as $field) {
                $orderIdentity[] = stockKey($record[$field]);
            }
            $orderKey = json_encode($orderIdentity, JSON_THROW_ON_ERROR);
            if (!isset($orders[$orderKey])) {
                $orders[$orderKey] = ['record' => $record, 'lines' => []];
            }
            $orders[$orderKey]['lines'][] = [
                'match' => $match,
                'quantity' => (int) $quantity,
                'unit_price' => $unitPrice,
                'description' => $get('SKU DESCRIPTION'),
                'sku' => $get('CLIENT SKU NO.'),
                'source_row' => $row,
                'receiver_name' => $get('RECEIVER NAME'),
                'review_reason' => $reviewReason,
            ];
        }

        foreach ($orders as $order) {
            $record = $order['record'];
            foreach (stockOrderRows($order['lines']) as $package) {
                if ($outputRow > 20001) {
                    throw new RuntimeException('The converted workbook exceeds 20,000 output rows.');
                }
                $reviewLines = $package['reviews'] ?? [];
                $reviewReason = $reviewLines[0]['review_reason'] ?? null;
                // Repeated loose items can share an output row; retain every source row for review.
                foreach ($reviewLines as $reviewLine) {
                    $reviewRows[] = [
                        $reviewLine['source_row'], $outputRow, $reviewLine['sku'],
                        $reviewLine['description'], $reviewLine['quantity'],
                        $reviewLine['unit_price'] ?? '', $reviewLine['receiver_name'], $reviewLine['review_reason'],
                        'Kept unsplit; repeated identical items are combined within the order. Confirm the SKU/name and carton size.',
                    ];
                }
                // EDIT: FINAL EXCEL VALUES — these assignments control every exported row.
                // $package is calculated in inventory.php: stockPackRows() or stockOrderRows().
                // Change the formulas there for specific order types; override values here for all rows.
                $record['Item Description'] = $package['description']; // Final text in Item Description.
                $record['Quantity'] = $package['quantity'];
                $record['Unit Price'] = $package['unit_price'];
                // Quantity is pieces per row (e.g. 10 or 14); Unit Price is the price per piece.
                // Other final overrides can go here, e.g. $record['Delivery Instructions'] = 'Call on arrival';

                foreach ($templateHeaders as $index => $header) {
                    $address = Coordinate::stringFromColumnIndex($index + 1) . $outputRow;
                    $value = $record[$header];
                    if (in_array($header, $numericHeaders, true) && $value !== '') {
                        $target->setCellValueExplicit(
                            $address,
                            readNumber((string) $value, $header, $row),
                            DataType::TYPE_NUMERIC
                        );
                    } else {
                        // Preserve identifiers and prevent text becoming formulas.
                        $target->setCellValueExplicit($address, (string) $value, DataType::TYPE_STRING);
                    }
                }
                if ($reviewReason !== null) {
                    $target->getStyle("A{$outputRow}:AB{$outputRow}")->getFill()
                        ->setFillType('solid')->getStartColor()->setARGB('FFFFF2CC');
                }
                $outputRow++;
            }
        }

        if ($outputRow === 2) {
            throw new RuntimeException('The source Excel has no data rows.');
        }

        $finalRow = $outputRow - 1;

        $target->freezePane('A2');
        $target->setAutoFilter("A1:AB{$finalRow}");
        $target->getStyle('A1:AB1')->getFont()->setBold(true);
        $target->getStyle('A1:AB1')->getFill()
            ->setFillType('solid')
            ->getStartColor()->setARGB('FFDCE6F1');
        $target->getStyle('A1:AB1')->getAlignment()->setWrapText(true);
        $target->getRowDimension(1)->setRowHeight(32);

        foreach (range(1, 28) as $column) {
            $target->getColumnDimension(
                Coordinate::stringFromColumnIndex($column)
            )->setWidth(22);
        }

        $target->getColumnDimension('J')->setWidth(50);
        $target->getColumnDimension('O')->setWidth(50);
        $target->getColumnDimension('W')->setWidth(60);
        $target->getStyle("W2:W{$finalRow}")->getAlignment()->setWrapText(true);
        $target->getStyle("V2:V{$finalRow}")
            ->getNumberFormat()->setFormatCode('0.00');
        // EDIT: PRICE DISPLAY — controls decimal formatting only; calculation is in inventory.php.
        $target->getStyle("Z2:Z{$finalRow}")
            ->getNumberFormat()->setFormatCode('0.00########');

        if ($reviewRows) {
            $review = $output->createSheet();
            $review->setTitle('Needs review');
            $reviewHeaders = [
                'Source Row', 'Template Row', 'Client SKU', 'Original Item Description',
                'Original Order Qty', 'Original Unit Price', 'Receiver Name', 'Reason', 'Action',
            ];
            foreach (array_merge([$reviewHeaders], $reviewRows) as $index => $values) {
                foreach ($values as $column => $value) {
                    $numeric = $index > 0 && in_array($column, [0, 1, 4, 5], true) && $value !== '';
                    $review->setCellValueExplicit(
                        Coordinate::stringFromColumnIndex($column + 1) . ($index + 1),
                        $value,
                        $numeric ? DataType::TYPE_NUMERIC : DataType::TYPE_STRING
                    );
                }
            }
            $lastReviewRow = count($reviewRows) + 1;
            $review->freezePane('A2');
            $review->setAutoFilter("A1:I{$lastReviewRow}");
            $review->getStyle('A1:I1')->getFont()->setBold(true);
            $review->getStyle('A1:I1')->getFill()->setFillType('solid')
                ->getStartColor()->setARGB('FFFFF2CC');
            $review->getStyle("A1:I{$lastReviewRow}")->getAlignment()->setWrapText(true);
            foreach (['A' => 12, 'B' => 14, 'C' => 24, 'D' => 60, 'E' => 18,
                      'F' => 18, 'G' => 28, 'H' => 65, 'I' => 65] as $column => $width) {
                $review->getColumnDimension($column)->setWidth($width);
            }
            $review->getStyle("F2:F{$lastReviewRow}")->getNumberFormat()->setFormatCode('0.00');
            $output->setActiveSheetIndex(0);
        }

        // Finish generating the workbook before starting the download.
        $downloadPath = tempnam(sys_get_temp_dir(), 'courier_');

        if ($downloadPath === false) {
            throw new RuntimeException('Unable to create the download file.');
        }

        $writer = new Xlsx($output);
        $writer->setPreCalculateFormulas(false);
        $writer->save($downloadPath);

        header(
            'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        header(
            'Content-Disposition: attachment; filename="courier_template_'
            . date('Ymd_His') . '.xlsx"'
        );
        header('Cache-Control: no-store');
        header('Content-Length: ' . filesize($downloadPath));

        readfile($downloadPath);
    } catch (Throwable $exception) {
        $error = $exception->getMessage();
    } finally {
        if (is_string($downloadPath) && is_file($downloadPath)) {
            unlink($downloadPath);
        }

        $input?->disconnectWorksheets();
        $output?->disconnectWorksheets();
    }

    if ($error === '') {
        exit;
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>SG Courier Excel Converter</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f3f5f8;
            padding: 40px 16px;
        }
        main {
            max-width: 540px;
            margin: auto;
            background: white;
            padding: 28px;
            border-radius: 12px;
        }
        h1 { font-size: 24px; }
        p { color: #475569; line-height: 1.5; }
        input { display: block; margin: 20px 0; }
        button {
            background: #2563eb;
            color: white;
            border: 0;
            border-radius: 6px;
            padding: 12px 18px;
            cursor: pointer;
        }
        .error {
            color: #991b1b;
            background: #fee2e2;
            padding: 12px;
            border-radius: 6px;
        }
    </style>
</head>
<body>
<main>
    <h1>SG Courier Excel Converter</h1>

    <?php if ($error !== ''): ?>
        <p class="error" role="alert"><?= escapeHtml($error) ?></p>
    <?php endif; ?>

    <form method="post" enctype="multipart/form-data">
        <label for="excel">Source Excel file (.xlsx or .xls, maximum 10 MB)</label>
        <input
            type="file"
            id="excel"
            name="excel"
            accept=".xlsx,.xls"
            required
        >
        <button type="submit">Convert and Download</button>
    </form>
</main>
</body>
</html>
