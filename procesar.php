<?php

require 'vendor/autoload.php'; // Incluye el autoloader de Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// --- CONFIGURACIÓN ---
$excelFilePath = 'Motos.xlsx'; 
$sheetName = 'Inscripciones';
$startRow = 9; // Última fila de encabezados
$startColumn = 'B'; 

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
        /* Estilos adicionales para centrar el contenedor en la vista */
        html, body { height: 100%; }
        body { display: flex; justify-content: center; align-items: center; }
    </style>
</head>
<body>
    <div class="container"> <h1 class="text-center">Resultado del Trámite</h1>
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

// Comprobación si se enviaron datos por POST
if ($_SERVER["REQUEST_METHOD"] == "POST") {

    try {
        // 1. Cargar el archivo de Excel existente
        $spreadsheet = IOFactory::load($excelFilePath);
        $worksheet = $spreadsheet->getSheetByName($sheetName);
        $configWorksheet = $spreadsheet->getSheetByName('Config');

        if (!$worksheet || !$configWorksheet) {
            $error_message = "La hoja de cálculo '{$sheetName}' o 'Config' no fue encontrada.";
            echo generate_result_page("Error de Archivo", $error_message, false);
            exit(); 
        }

        // Obtener el tipo de operación para decidir la lógica
        $tipo_operacion = $_POST['tipo_operacion'] ?? 'Inicial'; 

        // --- INICIALIZACIÓN DE VARIABLES DE CÁLCULO ---
        
        $caja = 0;  // Valor de Caja (Columna S)
        $cupl_final = 0; // Valor de CUPL (Columna T)
        $tipo_tramite = ''; // Q
        $transito_final = ''; // R

        // ✅ OBTENER COMISIÓN DE INSCRIPCIÓN INICIAL (N3)
        $valor_config_n3 = (float)$configWorksheet->getCell('N3')->getCalculatedValue();
        
        // ✅ NUEVA VARIABLE CONDICIONAL: Inicialmente asignada a la comisión de Inscripción Inicial
        $COMI = 0; // Inicializar
        
        if ($tipo_operacion === 'Inicial') {
            // Lógica para INSCRIPCIÓN INICIAL
            
            $tipo_pago = $_POST['COC'] ?? '';
            $transito_seleccionado = $_POST['transito'] ?? '';

            $tipo_tramite = $tipo_pago;
            $transito_final = $transito_seleccionado;

            // --- CUPL (Valor de Credito/Contado: D5, E5) ---
            if ($tipo_pago === 'Credito') { 
                $cupl_final = $configWorksheet->getCell('D5')->getCalculatedValue(); 
            } elseif ($tipo_pago === 'Contado') {
                $cupl_final = $configWorksheet->getCell('E5')->getCalculatedValue(); 
            }
            $cupl_final = (int)$cupl_final; // Columna T

            // --- CAJA (Valor de inscripción inicial: I5, J5, K5) ---
            switch ($transito_seleccionado) {
                case 'Villa':
                    $caja = $configWorksheet->getCell('I5')->getCalculatedValue(); 
                    break;
                case 'Cucuta':
                    $caja = $configWorksheet->getCell('J5')->getCalculatedValue(); 
                    break;
                case 'Patios':
                    $caja = $configWorksheet->getCell('K5')->getCalculatedValue(); 
                    break;
                default:
                    $caja = 0;
            }
            $caja = (int)$caja; // Columna S
            
            // ✅ ASIGNAR COMISIÓN N3 A $COMI
            $COMI = $valor_config_n3;


        } elseif ($tipo_operacion === 'Otro') {
            // Lógica para OTROS TRÁMITES
            
            $tramite_row = (int)($_POST['tipo_tramite_otro'] ?? 0); 
            $transito_otro = $_POST['transito_otro'] ?? '';
            
            if ($tramite_row >= 12) {
                
                $tipo_tramite = $configWorksheet->getCell('D' . $tramite_row)->getValue(); 
                $transito_final = $transito_otro; 

                // Offset de columnas para Cúcuta(0, E, F), Villa(2, G, H), Patios(4, I, J)
                $colOffset = 0; 
                if ($transito_otro === 'Villa') {
                    $colOffset = 2; 
                } elseif ($transito_otro === 'Patios') {
                    $colOffset = 4;
                }

                // Cálculo de Caja (E, G, I) y CUPL (F, H, J) - Basado en la tabla de trámites (fila 12 en adelante)
                $cajaCol = chr(ord('E') + $colOffset); 
                $cuplCol = chr(ord('F') + $colOffset); 

                $caja = (int)$configWorksheet->getCell($cajaCol . $tramite_row)->getCalculatedValue(); // Columna S
                $cupl_final = (int)$configWorksheet->getCell($cuplCol . $tramite_row)->getCalculatedValue(); // Columna T

                // ✅ OBTENER COMISIÓN DE LA COLUMNA K
                $comision_otro_k = (float)$configWorksheet->getCell('K' . $tramite_row)->getCalculatedValue();
                
                // ✅ ASIGNAR COMISIÓN K A $COMI
                $COMI = $comision_otro_k;
            }
        }
        
        // --- CÁLCULO DE IMPUESTO (CONDICIONAL) ---
        $impuesto = 0.0; // Inicializar a 0.0
        $base_imponible = (float)($_POST['BI'] ?? 0);
        $cilindraje = (float)($_POST['cilindraje'] ?? 0);

        // El cálculo de impuesto solo debe ocurrir si el cilindraje es >= 126
        if($cilindraje >= 126){
            $mes_actual = (int)date('n'); 
            $meses_restantes = 12 - $mes_actual + 1;
            $resultado = (($base_imponible * 0.015) / 12);
            $impuesto = round(($resultado * $meses_restantes) + 47500); 
        }

        // --- CÁLCULO DE TOTAL FACTURA ---
        // Total = Caja (S) + CUPL (T) + Impuesto (P) + Costo Base (N5) + Comisión Condicional (COMI)
        $total_factura = $caja + $cupl_final + $impuesto + $COMI;


        // ----------------------------------------------------------------------
        // --- CÁLCULO DEL ID DE TRÁMITE Y FILA DE INSERCIÓN ---
        // ----------------------------------------------------------------------
        
        $highestTrámiteID = 0;
        $lastDataRow = $startRow; 
        $highestRow = $worksheet->getHighestRow();
        
        for ($row = $startRow; $row <= $highestRow; $row++) {
            $cellValue = $worksheet->getCell('A' . $row)->getCalculatedValue();
            $trámiteID = (int)$cellValue; 
            
            if ($trámiteID > $highestTrámiteID) {
                $highestTrámiteID = $trámiteID; 
                $lastDataRow = $row; 
            }
        }
        
        $currentRow = $lastDataRow + 1; 

        // EVITAR SOBREESCRIBIR CORTES CONSECUTIVOS
        while ($worksheet->getCell('A' . $currentRow)->getValue() === 'CORTE') {
            $currentRow++;
        }
        
        $indice_numerico = $highestTrámiteID + 1;
        $worksheet->setCellValue('A' . $currentRow, $indice_numerico); 
        
if ($tipo_operacion === 'Inicial') {
    // Si es "Inicial", la columna W debe decir "Inscripción Inicial"
    $nombre_tramite_col_W = 'Inscripción Inicial';
} else {
    // Si es "Otro", la columna W debe llevar el nombre específico
    $nombre_tramite_col_W = $tipo_tramite;
}

        // 3. Insertar los datos en la nueva fila, comenzando en la columna B
        // ***************************************************************
        // ESTE ARRAY SE MANTIENE SIN CAMBIOS, usando $total_factura (que ya incluye $COMI)
        // ***************************************************************
        $data = [
            // Cliente: B, C, D, E, F, G 
            $_POST['nombres'] ?? '',     // B
            $_POST['apellidos'] ?? '',   // C
            $_POST['cedula'] ?? '',      // D
            $_POST['direccion'] ?? '',   // E
            $_POST['telefono'] ?? '',    // F
            $_POST['no_fac'] ?? '',      // G (Número de Factura)

            // Moto: H, I, J, K, L, M 
            $_POST['placa'] ?? '',       // H
            $_POST['marca'] ?? '',       // I
            $_POST['linea'] ?? '',       // J
            $_POST['no_motor'] ?? '',    // K
            $_POST['no_chasis'] ?? '',   // L
            $_POST['ano'] ?? '',         // M

            // Cálculos y Precio: N, O, P, Q, R, S, T, U, V
            (int)$cilindraje,            // N (Cilindraje)
            $base_imponible,             // O (Precio Base/BI)
            $impuesto,                   // P (Impuesto)
            $tipo_tramite,               // Q (Tipo: Contado/Credito o Nombre del Trámite)
            $transito_final,             // R (Tránsito: Villa/Cucuta/Patios)
            $caja,                       // S (Caja)
            $cupl_final,                 // T (CUPL)
            date('Y-m-d'),
            $COMI ?? '',
            $nombre_tramite_col_W ?? '',               // U (Fecha)
            $total_factura               // V (Total)
        ];
        
        $currentColumn = $startColumn;
        foreach ($data as $value) {
            $worksheet->setCellValue($currentColumn . $currentRow, $value);
            $currentColumn++; 
        }
        
        // 6. Guardar el archivo de Excel modificado
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($excelFilePath);

        // 7. Mensaje de éxito
        $success_message = "Datos insertados correctamente en la fila {$currentRow} con el ID #{$indice_numerico}.";
        echo generate_result_page("Éxito de Inserción", $success_message, true);

    } catch (\Exception $e) {
        $error_message = "Ocurrió un error: " . $e->getMessage();
        echo generate_result_page("Error", $error_message, false);
    }
} else {
    // Si no se accedió mediante POST
    header("Location: Index.php");
    exit();
}

?>