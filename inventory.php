<?php
declare(strict_types=1);

class StockItemReviewException extends RuntimeException {}

function stockKey(string $value): string
{
    $value = str_replace("\xC2\xA0", ' ', $value);
    return mb_strtoupper(trim(preg_replace('/\\s+/u', ' ', $value) ?? $value), 'UTF-8');
}

function stockPositiveInteger(string $value, string $label): int
{
    if (!preg_match('/^[0-9]+$/D', $value)
        || strlen($value) > 9 || (int) $value < 1) {
        throw new RuntimeException($label . ' must be a positive whole number (maximum 999999999).');
    }
    return (int) $value;
}

function stockProductsFromWorkbook(string $path): array
{
    $type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($path, ['Xlsx', 'Xls']);
    $book = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($type)->load($path);
    try {
        $sheet = $book->getSheet(0);
        if ($sheet->getHighestDataRow() > 10000) {
            throw new RuntimeException('The product workbook exceeds 10,000 rows.');
        }
        $read = static function (string $address) use ($sheet): string {
            $cell = $sheet->getCell($address);
            if ($cell->getDataType() === \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
                throw new RuntimeException('Paste the value of formula cell ' . $address . ' before importing.');
            }
            return trim((string) $cell->getValue());
        };
        foreach (['C6' => 'SKU#', 'D6' => 'NAME', 'F6' => 'SKU# FOR CARTON',
                  'G6' => 'SKU# FOR LOOSE', 'H6' => 'CARTON'] as $address => $expected) {
            if (stockKey($read($address)) !== $expected) {
                throw new RuntimeException('Use the StockCount_FinishedGoods layout; unexpected header at ' . $address);
            }
        }
        $products = [];
        for ($row = 7; $row <= $sheet->getHighestDataRow(); $row++) {
            $sku = stockKey($read('C' . $row));
            if ($sku === '' || $sku === '.') { continue; }
            if (isset($products[$sku])) {
                throw new RuntimeException('Duplicate product SKU on row ' . $row . ': ' . $sku);
            }
            $name = $read('D' . $row);
            if ($name === '' || mb_strlen($name) > 255 || mb_strlen($sku) > 100) {
                throw new RuntimeException('Invalid product name or SKU on row ' . $row);
            }
            $product = [
                'sku' => $sku,
                'material_class' => $read('B' . $row) ?: null,
                'product_name' => $name,
                'abbreviation' => $read('E' . $row) ?: null,
                'carton_sku' => stockKey($read('F' . $row)) ?: null,
                'loose_sku' => stockKey($read('G' . $row)) ?: null,
                'quantity_per_carton' => stockPositiveInteger($read('H' . $row), 'Carton quantity on row ' . $row),
            ];
            foreach (['material_class', 'abbreviation', 'carton_sku', 'loose_sku'] as $key) {
                if (mb_strlen($product[$key] ?? '') > 100) {
                    throw new RuntimeException('Value too long on row ' . $row . ': ' . $key);
                }
            }
            $products[$sku] = $product;
        }
        if (!$products) { throw new RuntimeException('No product rows were found.'); }
        stockAliases($products);
        return $products;
    } finally {
        $book->disconnectWorksheets();
    }
}

/* Only explicitly recorded carton/loose SKUs are recognised. */
function stockAliases(array $products): array
{
    $aliases = [];
    foreach ($products as $product) {
        foreach (['carton_sku', 'loose_sku'] as $kind) {
            $sku = stockKey((string) ($product[$kind] ?? ''));
            if ($sku === '') { continue; }
            if (isset($aliases[$sku])) {
                throw new RuntimeException('Carton/loose SKU is assigned more than once: ' . $sku);
            }
            $aliases[$sku] = [
                'product' => $product,
                'boxes_per_order_unit' => $kind === 'carton_sku'
                    ? (int) $product['quantity_per_carton'] : 1,
            ];
        }
    }
    return $aliases;
}

function stockImport(PDO $db, array $products): int
{
    $db->beginTransaction();
    try {
        // Use the current table's columns; optional stock workbook fields may have been removed.
        $metadata = $db->query('SELECT * FROM products WHERE 1=0');
        $available = [];
        for ($column = 0; $column < $metadata->columnCount(); $column++) {
            $available[] = $metadata->getColumnMeta($column)['name'];
        }
        $columns = array_values(array_intersect(
            ['material_class', 'product_name', 'abbreviation', 'carton_sku', 'loose_sku', 'quantity_per_carton'],
            $available
        ));
        // Update only by primary SKU; unique alias collisions must fail, not update another product.
        $update = $db->prepare('UPDATE products SET '
            . implode(', ', array_map(static fn(string $column): string => $column . '=?', $columns))
            . ' WHERE sku=?');
        $exists = $db->prepare('SELECT sku FROM products WHERE sku=?');
        $insert = $db->prepare('INSERT INTO products (sku, ' . implode(', ', $columns) . ') VALUES ('
            . implode(', ', array_fill(0, count($columns) + 1, '?')) . ')');
        foreach ($products as $p) {
            $fields = array_map(static fn(string $column) => $p[$column] ?? null, $columns);
            $exists->execute([$p['sku']]);
            if ($exists->fetchColumn() !== false) {
                $update->execute([...$fields, $p['sku']]);
            } else {
                $insert->execute([$p['sku'], ...$fields]);
            }
        }
        stockLoadCatalog($db); // Validate cross-column aliases before committing.
        $db->commit();
        return count($products);
    } catch (Throwable $error) {
        if ($db->inTransaction()) { $db->rollBack(); }
        throw $error;
    }
}

function stockLoadCatalog(PDO $db): array
{
    $products = [];
    foreach ($db->query('SELECT * FROM products') as $product) {
        $products[stockKey($product['sku'])] = $product;
    }
    if (!$products) { throw new RuntimeException('Import the stock workbook first.'); }
    $rules = ['sku' => [], 'description' => []];
    foreach ($db->query('SELECT * FROM order_item_rules') as $rule) {
        $product = $products[stockKey($rule['product_sku'])] ?? null;
        if (!$product) { throw new RuntimeException('Rule points to a missing product.'); }
        $key = stockKey($rule['match_value']);
        if (isset($rules[$rule['match_type']][$key])) {
            throw new RuntimeException('Duplicate order-item rule: ' . $key);
        }
        $rules[$rule['match_type']][$key] = [
            'product' => $product,
            'boxes_per_order_unit' => (int) $rule['boxes_per_order_unit'],
        ];
    }
    return ['products' => $products, 'aliases' => stockAliases($products), 'rules' => $rules];
}

function stockNameKey(string $value): string
{
    return trim(preg_replace('/[^\p{L}\p{N}]+/u', ' ', stockKey($value)) ?? '');
}

/* Read advertised retail quantities, not product sizes such as 20 SACHETS or 1KG. */
function stockDescriptionBoxes(string $description, int $cartonSize, int $row): ?int
{
    $description = stockKey($description);
    $units = 'CARTONS?|CTNS?|BOXES|BOX|TUBS?|JARS?|BOTTLES?|PCS|PIECES?|UNITS?';
    $term = '\d{1,9}\s*(?:(?:' . $units . ')\b\s*)?(?:(?:FREE|BONUS)\b\s*)?';
    $expression = '/(?<![\p{L}\p{N}.])' . $term
        . '(?:(?:\+|\bPLUS\b)\s*(?:(?:FREE|BONUS)\s*)?' . $term . ')+'
        . '(?![\p{L}\p{N}.])/u';
    preg_match_all($expression, $description, $sums);
    if (count($sums[0]) > 1) {
        throw new StockItemReviewException("Row {$row}: multiple bundle quantities were found in the description.");
    }
    if ($sums[0]) {
        $parts = preg_split('/\+|\bPLUS\b/u', $sums[0][0]);
        $amounts = [];
        $lastUnit = '';
        foreach ($parts as $part) {
            $part = trim(preg_replace('/\b(?:FREE|BONUS)\s*/u', '', $part) ?? '');
            if (!preg_match('/^(\d+)\s*(' . $units . ')?\s*$/u', $part, $m)) {
                throw new StockItemReviewException("Row {$row}: unsupported bundle quantity.");
            }
            $unit = $m[2] ?? '';
            if ($unit !== '') { $lastUnit = $unit; }
            $amounts[] = [(int) $m[1], $unit];
        }
        $total = 0;
        foreach ($amounts as [$amount, $unit]) {
            $unit = $unit !== '' ? $unit : $lastUnit;
            $total += $amount * (preg_match('/^(?:CARTON|CTN)/', $unit) ? $cartonSize : 1);
        }
        if ($total < 1) { throw new StockItemReviewException("Row {$row}: bundle quantity must be positive."); }
        return $total;
    }
    preg_match_all('/(?<![\p{L}\p{N}.])(\d{1,9})\s*(' . $units . ')\b/u', $description, $counts, PREG_SET_ORDER);
    $quantities = [];
    foreach ($counts as $count) {
        $quantities[] = (int) $count[1]
            * (preg_match('/^(?:CARTON|CTN)/', $count[2]) ? $cartonSize : 1);
    }
    $quantities = array_values(array_unique($quantities));
    if (count($quantities) > 1) {
        throw new StockItemReviewException("Row {$row}: conflicting box/carton quantities were found in the description.");
    }
    if ($quantities && $quantities[0] < 1) {
        throw new StockItemReviewException("Row {$row}: box quantity must be positive.");
    }
    return $quantities[0] ?? null;
}

function stockResolve(array $catalog, string $sku, string $description, int $row): array
{
    $sku = stockKey($sku);
    $description = stockKey($description);
    // Preserve any explicit overrides already saved by the user.
    $match = $catalog['rules']['sku'][$sku]
        ?? $catalog['rules']['description'][$description]
        ?? null;
    if ($match) { return $match; }

    $match = $catalog['aliases'][$sku] ?? null;
    if (!$match && isset($catalog['products'][$sku])) {
        $match = ['product' => $catalog['products'][$sku], 'boxes_per_order_unit' => 1];
    }
    if (!$match) {
        $name = ' ' . stockNameKey($description) . ' ';
        $candidates = [];
        foreach ($catalog['products'] as $product) {
            $productName = stockNameKey($product['product_name']);
            if ($productName !== '' && str_contains($name, ' ' . $productName . ' ')) {
                $candidates[] = ['name' => $productName, 'product' => $product];
            }
        }
        usort($candidates, static fn(array $a, array $b): int => strlen($b['name']) <=> strlen($a['name']));
        if ($candidates) {
            $best = $candidates[0];
            foreach (array_slice($candidates, 1) as $candidate) {
                if (!str_contains(' ' . $best['name'] . ' ', ' ' . $candidate['name'] . ' ')
                    || $best['name'] === $candidate['name']) {
                    throw new StockItemReviewException("Row {$row}: the description matches more than one product; supply its SKU.");
                }
            }
            $match = ['product' => $best['product'], 'boxes_per_order_unit' => 1];
        }
    }
    if (!$match) {
        throw new StockItemReviewException("Row {$row}: product not found in the database: {$sku} / {$description}.");
    }
    $boxes = stockDescriptionBoxes($description, (int) $match['product']['quantity_per_carton'], $row);
    if ($boxes !== null) { $match['boxes_per_order_unit'] = $boxes; }
    return $match;
}

/*
 * One output row represents one full carton-size pack or a final partial pack.
 * Quantity is pieces per row; Unit Price is per piece to preserve the source total.
 */
function stockPackRows(array $match, int $orderQuantity, ?float $unitPrice): Generator
{
    $factor = (int) $match['boxes_per_order_unit'];
    $size = (int) $match['product']['quantity_per_carton'];
    if ($factor < 1 || $size < 1 || $orderQuantity < 1 || $orderQuantity > intdiv(1000000, $factor)) {
        throw new RuntimeException('Invalid pack size or more than 1,000,000 boxes in one order row.');
    }
    $totalBoxes = $orderQuantity * $factor;
    $rowCount = (int) ceil($totalBoxes / $size);
    if ($rowCount > 20000) {
        throw new RuntimeException('One order row would produce more than 20,000 output rows.');
    }
    $totalCents = null;
    if ($unitPrice !== null) {
        $amount = $unitPrice * $orderQuantity * 100;
        if (!is_finite($amount) || $amount < 0 || $amount > 1000000000000) {
            throw new RuntimeException('Order value exceeds the supported range.');
        }
        $totalCents = (int) round($amount);
    }
    $allocated = 0;
    for ($boxesDone = 0; $boxesDone < $totalBoxes;) {
        $boxes = min($size, $totalBoxes - $boxesDone);
        $boxesDone += $boxes;
        $price = '';
        if ($totalCents !== null) {
            $cumulative = (int) round($totalCents * ($boxesDone / $totalBoxes));
            $price = ($cumulative - $allocated) / (100 * $boxes);
            $allocated = $cumulative;
        }
        yield [
            'boxes' => $boxes,
            'description' => $match['product']['product_name'] . ' (' . $boxes
                . ($boxes === 1 ? ' Box)' : ' Boxes)'),
            'quantity' => $boxes,
            'unit_price' => $price,
        ];
    }
}

/* EDIT: SAME-ITEM GROUPING — different products retain separate descriptions within each order. */
function stockOrderRows(array $lines): Generator
{
    $items = [];
    foreach ($lines as $line) {
        $match = $line['match'];
        $key = $match !== null
            ? 'product:' . stockKey($match['product']['sku'])
            : 'loose:' . json_encode([stockKey($line['sku']), stockKey($line['description'])], JSON_THROW_ON_ERROR);
        if (!isset($items[$key])) {
            $items[$key] = [
                'match' => $match, 'description' => $line['description'],
                'pieces' => 0, 'cents' => 0, 'priced' => true, 'reviews' => [],
            ];
        }
        $factor = $match === null ? 1 : (int) $match['boxes_per_order_unit'];
        if ($factor < 1 || $line['quantity'] < 1
            || $line['quantity'] > intdiv(999999999, $factor)) {
            throw new RuntimeException('Item quantity exceeds the supported range.');
        }
        $items[$key]['pieces'] += $factor * $line['quantity'];
        if ($items[$key]['pieces'] > 999999999) {
            throw new RuntimeException('Combined item quantity exceeds the supported range.');
        }
        if ($line['unit_price'] === null) {
            $items[$key]['priced'] = false;
        } else {
            $amount = $line['unit_price'] * $line['quantity'] * 100;
            if (!is_finite($amount) || $amount < 0 || $amount > 1000000000000) {
                throw new RuntimeException('Order value exceeds the supported range.');
            }
            $items[$key]['cents'] += (int) round($amount);
            if ($items[$key]['cents'] > 1000000000000) {
                throw new RuntimeException('Combined item value exceeds the supported range.');
            }
        }
        if ($match === null) { $items[$key]['reviews'][] = $line; }
    }
    foreach ($items as $item) {
        $perPiece = $item['priced'] ? $item['cents'] / (100 * $item['pieces']) : null;
        $match = $item['match'];
        if ($match === null || (int) $match['product']['quantity_per_carton'] <= 1) {
            // EDIT: LOOSE ITEM VALUES — retain the source name; 1 + 1 + 1 becomes Quantity 3.
            yield [
                'description' => $item['description'],
                'quantity' => $item['pieces'],
                'unit_price' => $perPiece ?? '',
                'reviews' => $item['reviews'],
            ];
        } else {
            // EDIT: CARTON DESCRIPTION — retain the earlier PRODUCT NAME (10 Boxes) format.
            yield from stockPackRows(
                ['product' => $match['product'], 'boxes_per_order_unit' => 1],
                $item['pieces'],
                $perPiece
            );
        }
    }
}

