<?php

require 'PhpSpreadsheet-master/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Load the invoice template Excel file
$invoice_template = IOFactory::load('invoice_template.xlsx');

// Select the sheet where you want to insert data
$worksheet = $invoice_template->getSheetByName('Sheet1');

// Define the product list
$product_list = [];

$csvfile = fopen('product_list.csv', 'r');
while (($row = fgetcsv($csvfile)) !== false) {
    $product_list[] = [
        'name' => $row[0],
        'number' => $row[1],
        'hs' => $row[2],
        'packing' => $row[3]
    ];
}

fclose($csvfile);

// Get the selected products
$selected_products = [];

foreach ($worksheet->getRange('A23:B42') as $row) {
    $product_id = $row['A'];
    $product_name = $row['B'];
    $product_quantity = $row['F'];
    $product_price = $row['G'];

    $product = array_filter(array_merge($product_list[intval($product_id) - 1], [
        'quantity' => $product_quantity,
        'price' => $product_price,
        'total_price' => $product_quantity * $product_price
    ]));

    $selected_products[] = $product;
}

// Calculate the total price
$total_price = array_sum(array_column($selected_products, 'total_price'));

// Write the invoice data to the Excel file
$worksheet->getRange('A13')->setValue($invoice_number);
$worksheet->getRange('A14')->setValue($invoice_date);
$worksheet->getRange('A15')->setValue($lpo_number);
$worksheet->getRange('E15')->setValue($lpo_date);
$worksheet->getRange('A17')->setValue($company_name);
$worksheet->getRange('A18')->setValue($customer_name);
$worksheet->getRange('A19')->setValue($cr_number);
$worksheet->getRange('A20')->setValue($vatreg_number);
$worksheet->getRange('J14')->setValue($salesman);
$worksheet->getRange('J17')->setValue($contact_person);
$worksheet->getRange('J18')->setValue($designation);
$worksheet->getRange('J19')->setValue($contact_number);
$worksheet->getRange('J20')->setValue($omanID);

foreach ($selected_products as $product) {
    $row = $product['serial_number'] + 23;
    $worksheet->getRange('A' . $row)->setValue($product['serial_number']);
    $worksheet->getRange('B' . $row)->setValue($product['number']);
    $worksheet->getRange('C' . $row)->setValue($product['name']);
    $worksheet->getRange('D' . $row)->setValue($product['hs']);
    $worksheet->getRange('E' . $row)->setValue($product['packing']);
    $worksheet->getRange('F' . $row)->setValue($product['quantity']);
    $worksheet->getRange('G' . $row)->setValue($product['price']);
    $worksheet->getRange('H' . $row)->setValue($product['total_price']);
    $worksheet->getRange('I' . $row)->setValue($product['vat']);
    $worksheet->getRange('J' . $row)->setValue($product['sum_price']);
    $worksheet->getRange('I42' . $row)->setValue($product['total_price']);
    $worksheet->getRange('J42' . $row)->setValue($product['vat']);
}

$worksheet->getRange('J36')->setValue($total_price);

// Save the invoice file
$writer = new Xlsx($invoice_template);
$writer->save('invoice-' . $invoice_number . '.xlsx');


?>

