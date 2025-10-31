<?php

require 'vendor/autoload.php'; // Incluye el autoloader de Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// --- CONFIGURACIÓN ---
$excelFilePath = 'Motos.xlsx';
$sheetName = 'Inscripciones'; // Hoja donde se registra la fecha
$configSheetName = 'Config';   // Hoja para guardar el contador
$dateColumn = 'U';             // Columna de la Fecha/Corte en 'Inscripciones'
$counterCell = 'A1';           // Celda para guardar el número de semana en 'Config' (Se mantiene como respaldo)

// Función para generar la estructura HTML de la página de resultados
function generate_result_page($title, $message, $is_success) {
    $alert_class = $is_success ? 'alert-success' : 'alert-danger';
    $icon = $is_success ? '✅' : '❌';
    
    return <<<HTML
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{$title}</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="styles.css">
    <style>
        /* Estilos para centrar el contenedor */
        html, body { height: 100%; }
        body { display: flex; justify-content: center; align-items: center; }
    </style>
</head>
<body>
    <div class="container"> <h1 class="text-center">Proceso de Corte</h1>
        <div class="alert {$alert_class}" role="alert">
            <h4 class="alert-heading">{$icon} {$title}</h4>
            <p>{$message}</p>
        </div>
        <div class="text-center mt-4">
            <a href="Index.php" class="btn btn-primary btn-lg">Volver al Formulario</a>
        </div>
    </div>
</body>
</html>
HTML;
}

if ($_SERVER["REQUEST_METHOD"] == "POST" && ($_POST['corte_switch'] ?? '') === 'Si') {
    try {
        // 1. Cargar el archivo de Excel
        $spreadsheet = IOFactory::load($excelFilePath);
        $worksheet = $spreadsheet->getSheetByName($sheetName);
        $configWorksheet = $spreadsheet->getSheetByName($configSheetName);

        if (!$worksheet || !$configWorksheet) {
            $error_message = "Error: Hoja '{$sheetName}' o '{$configSheetName}' no encontrada.";
            echo generate_result_page("Error de Archivo", $error_message, false);
            exit(); 
        }
        
        // --- 2. LÓGICA PARA ENCONTRAR EL RANGO DE LA SUMA Y LA SEMANA MÁS ALTA ---
        $highestWeekFound = 0;
        $highestRow = $worksheet->getHighestRow();
        
        // La primera fila de datos es la 10
        $startDataRow = 10;
        // Inicializamos la última fila de corte válida a la fila anterior al inicio de los datos (9)
        $lastCutRow = $startDataRow - 1; 
        
        for ($row = $startDataRow; $row <= $highestRow; $row++) {
            // Buscamos el TAG 'CORTE' en la Columna A para encontrar el límite de la última suma
            if ($worksheet->getCell('A' . $row)->getValue() === 'CORTE') {
                $lastCutRow = $row;
            }
            
            // También buscamos el número de semana más alto en la columna de fecha (U)
            $cellValue = $worksheet->getCell($dateColumn . $row)->getValue();
            if (preg_match('/^Semana (\d+)$/i', $cellValue, $matches)) {
                $weekNumber = (int)$matches[1];
                if ($weekNumber > $highestWeekFound) {
                    $highestWeekFound = $weekNumber; 
                }
            }
        }

        // 3. Definir el rango de la suma
        // La suma comienza en la fila siguiente al último corte (o en la fila 10 si no hay cortes)
        $sumStartRow = $lastCutRow + 1;
        // La suma termina en la última fila con datos ($highestRow)
        $sumEndRow = $highestRow;
        
        $newWeek = $highestWeekFound + 1;
        $tag = "Semana " . $newWeek;

        // Actualizar el contador en la hoja 'Config:A1' 
        $configWorksheet->setCellValue($counterCell, $newWeek);
        
        // --- FIN LÓGICA DE CORTES Y RANGO ---

        // 4. CALCULAR LAS SUMAS SOLICITADAS (P, T, S) - **¡Solo en el rango definido!**
        $totalP = 0; // Columna P (Impuesto)
        $totalT = 0; // Columna T (Total)
        $totalS = 0; // Columna S (Valor Crédito/Contado)
        $totalV = 0; // ✅ Columna V (Total Comisiones)
        $totalX = 0;
        
        // Iteramos sobre las filas de datos desde $sumStartRow hasta $sumEndRow
        for ($row = $sumStartRow; $row <= $sumEndRow; $row++) {
            // Se usa getCalculatedValue() para obtener el valor numérico, importante para fórmulas
            
            // Suma de P (Impuesto)
            $valueP = $worksheet->getCell('P' . $row)->getCalculatedValue();
            if (is_numeric($valueP)) {
                $totalP += (float)$valueP;
            }

            // Suma de T (Total)
            $valueT = $worksheet->getCell('T' . $row)->getCalculatedValue();
            if (is_numeric($valueT)) {
                $totalT += (float)$valueT;
            }
            
            // Suma de S (Valor Crédito/Contado) - Solo si es numérico
            $valueS = $worksheet->getCell('S' . $row)->getCalculatedValue();
            if (is_numeric($valueS)) {
                $totalS += (float)$valueS;
            }
            // ✅ Suma V (Total Factura)
            $valV = $worksheet->getCell('V' . $row)->getCalculatedValue();
            if (is_numeric($valV)) { $totalV += (float)$valV; }

            // ✅ Suma X (Total Factura)
            $valX = $worksheet->getCell('X' . $row)->getCalculatedValue();
            if (is_numeric($valX)) { $totalX += (float)$valX; }
        }

        // 5. Encontrar la siguiente fila vacía en 'Inscripciones'
        // *** CORRECCIÓN IMPLEMENTADA: Buscamos la última fila con datos reales en Columna A ***
        $lastDataRow = 0;
        // Iteramos desde la fila de inicio de datos (10) hasta la más alta.
        for ($row = $startDataRow; $row <= $highestRow; $row++) {
            // Verificamos si la celda A tiene algún valor (el TAG 'CORTE' o un dato de inscripción)
            if (!empty($worksheet->getCell('A' . $row)->getValue())) {
                $lastDataRow = $row;
            }
        }
        
        // La fila de inserción es la siguiente a la última fila con datos reales
        $currentRow = $lastDataRow + 1; 

        // 6. Insertar el TAG de Corte y el indicador
        $worksheet->setCellValue($dateColumn . $currentRow, $tag);
        $worksheet->setCellValue('A' . $currentRow, 'CORTE');

        // 7. Insertar las sumas en la fila del corte (P, T, S, V)
        $worksheet->setCellValue('P' . $currentRow, $totalP);
        $worksheet->setCellValue('T' . $currentRow, $totalT);
        $worksheet->setCellValue('S' . $currentRow, $totalS); // Suma de S
        $worksheet->setCellValue('V' . $currentRow, $totalV);
        $worksheet->setCellValue('X' . $currentRow, $totalX);

        // 8. Guardar el archivo de Excel modificado
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($excelFilePath);

        $success_message = "Fecha de corte registrada como '{$tag}' en la fila {$currentRow}. 
        Se han registrado los totales: Total de Impuestos (P) = {$totalP}, Total de CUPL (T) = {$totalT}, Total de Caja (S) = {$totalS}, Total Comisiones (V) = {$totalV}, Total General (X) = {$totalX}.";
        echo generate_result_page("Corte Establecido", $success_message, true);

    } catch (\Exception $e) {
        $error_message = "Ocurrió un error al procesar el corte: " . $e->getMessage();
        echo generate_result_page("Error", $error_message, false);
    }
} else {
    // Si se accedió por POST pero la opción no fue 'Si' o no se accedió por POST
    $message = "No se estableció una fecha de corte o la opción seleccionada fue 'No'.";
    if ($_SERVER["REQUEST_METHOD"] == "POST") {
         echo generate_result_page("Proceso Omitido", $message, true);
    } else {
        header("Location: Index.php");
        exit();
    }
}

?>