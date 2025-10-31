<?php
require 'vendor/autoload.php'; // Incluye el autoloader de Composer

use PhpOffice\PhpSpreadsheet\IOFactory;

// --- CONFIGURACIÓN DE EXCEL ---
$excelFilePath = 'Motos.xlsx'; 
$sheetName = 'Inscripciones';
$startRow = 9; // Fila donde inician los datos (después de encabezados)

$cortes_data = [];

// Columnas de las sumas en la fila de CORTE (según tu estructura de Excel)
// 'S': Caja, 'T': CUPL, 'V': Comisión, 'X': Total
$SUM_COLUMNS = [
    'S' => 'Caja',
    'T' => 'CUPL',
    'V' => 'Comisión',
    'X' => 'Facturas'
];

try {
    // 1. Cargar el archivo de Excel
    $spreadsheet = IOFactory::load($excelFilePath);
    $worksheet = $spreadsheet->getSheetByName($sheetName);

    if (!$worksheet) {
        throw new Exception("La hoja de cálculo '{$sheetName}' no fue encontrada.");
    }

    $highestRow = $worksheet->getHighestRow();

    // 2. Recorrer los datos para encontrar las filas de CORTE
    for ($row = $startRow; $row <= $highestRow; $row++) {
        $cellA = $worksheet->getCell('A' . $row)->getCalculatedValue();

        // Verificar si la fila es una fila de CORTE
        if ($cellA === 'CORTE') {
            
            // Extraer el nombre de la semana (Columna U, según tu estructura)
            $semana = $worksheet->getCell('U' . $row)->getCalculatedValue();
            
            // Si la semana no tiene un nombre o está vacía, la ignoramos o le damos un nombre por defecto
            if (empty($semana)) {
                $semana = "Corte en Fila {$row}";
            }

            $corte = [
                'row_num' => $row,
                'semana' => $semana,
                'valores' => []
            ];

            // 3. Extraer los valores de las columnas sumadas
            foreach ($SUM_COLUMNS as $col_key => $col_name) {
                // Usamos getFormattedValue para obtener el valor tal como se ve en Excel (si tiene formato contable)
                $value = $worksheet->getCell($col_key . $row)->getFormattedValue();
                $corte['valores'][$col_name] = $value;
            }

            $cortes_data[] = $corte;
        }
    }
} catch (\Exception $e) {
    $error_message = "Error al leer el archivo Excel: " . $e->getMessage();
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cuenta de Cobro Semanal</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="styles.css">
    <style>
        .container {
            max-width: 900px;
            margin: 30px auto;
        }
        .cobro-header {
            background-color: #343a40;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
            margin-bottom: 0;
        }
        .cobro-detail {
            border: 1px solid #dee2e6;
            border-top: none;
            padding: 20px;
            margin-bottom: 30px;
        }
        .cobro-detail table th, .cobro-detail table td {
            font-size: 1.1em;
        }
        /* Estilo específico para impresión/PDF */
        @media print {
            .btn, .no-print {
                display: none !important;
            }
            .container {
                max-width: 100%;
                margin: 0;
                padding: 0;
            }
            body {
                background-color: white !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        
        <p class="text-center no-print">Selecciona la semana de corte para generar la cuenta de cobro.</p>

        <?php if (isset($error_message)): ?>
            <div class="alert alert-danger" role="alert"><?= $error_message ?></div>
        <?php endif; ?>

        <?php if (empty($cortes_data)): ?>
            <div class="alert alert-warning text-center">
                No se encontraron Fechas de CORTE en la hoja '<?= $sheetName ?>'.
            </div>
        <?php else: ?>
            
            <div class="form-group mb-4 no-print">
                <label for="selectCorte">Seleccionar Corte Semanal:</label>
                <select class="form-control" id="selectCorte">
                    <option value="">--- Seleccione una Semana ---</option>
                    <?php foreach ($cortes_data as $index => $corte): ?>
                        <option value="<?= $index ?>">
                            <?= htmlspecialchars($corte['semana']) ?> (Fila <?= $corte['row_num'] ?>)
                        </option>
                    <?php endforeach; ?>
                </select>
            </div>
            
            <button class="btn btn-primary btn-block mb-4 no-print" id="printButton" disabled>
                Generar PDF de la Cuenta de Cobro
            </button>
            
            <?php foreach ($cortes_data as $index => $corte): ?>
                <div class="cuenta-cobro-box" id="corte-<?= $index ?>" style="display: none;">
                    <div class="cobro-header text-center">
                        <h2>CUENTA DE COBRO SEMANAL</h2>
                    </div>
                    <div class="cobro-detail">
                        <p>
                            Fecha de Emisión: <?= date('Y-m-d') ?> <br>
                        </p>
                        
                        <h4 class="mt-4">Resumen de Totales</h4>
                        <table class="table table-bordered mt-3">
                            <thead class="thead-light">
                                <tr>
                                    <th>Concepto</th>
                                    <th>Valor Sumado</th>
                                </tr>
                            </thead>
                            <tbody>
                                <?php 
                                $total_neto_sum = 0; // Variable para la suma de todos los totales
                                $keys = array_keys($corte['valores']);
                                foreach ($keys as $key_name):
                                    $value_formatted = $corte['valores'][$key_name];
                                    
                                    // Limpiar el valor formateado para realizar la suma
                                    // Esto asume que el formato de moneda es solo '$' y ','
                                    $value_clean = str_replace(['$', ','], '', $value_formatted);
                                    
                                    // Sumamos el valor a la variable de total neto
                                    if (is_numeric($value_clean)) {
                                        $total_neto_sum += (float)$value_clean;
                                    }
                                ?>
                                    <tr>
                                        <td><?= htmlspecialchars($key_name) ?></td>
                                        <td><?= htmlspecialchars($value_formatted) ?></td>
                                    </tr>
                                <?php endforeach; ?>
                            </tbody>
                            <tfoot>
                                <tr class="table-dark">
                                    <td>TOTAL NETO</td>
                                    <td><?= '$' . number_format($total_neto_sum, 0) ?></td>
                                </tr>
                            </tfoot>
                        </table>
                        
                        <p class="text-left mt-4">Atentamente,</p>
                    <p class="text-center mt-4">Doris del Valle Diaz Cruz<br>NIT. 60.434.537-4</p>
                    <p class="text-center mt-4">Régimen Simplificado<br>Av. 13 # 43-80 LOS PATIOS<br>Tel. 3005144538</p>
                    </div>
                </div>
            <?php endforeach; ?>

        <?php endif; ?>
    </div>
    <div class="text-center">
<a href="registros.php" class="btn btn-secondary" >Volver al Formulario</a>
<p></p>

</div>
    <script>
        document.getElementById('selectCorte').addEventListener('change', function() {
            var selectedIndex = this.value;
            var printButton = document.getElementById('printButton');
            
            // Ocultar todos los cuadros de cuenta de cobro
            document.querySelectorAll('.cuenta-cobro-box').forEach(box => {
                box.style.display = 'none';
            });

            if (selectedIndex !== '') {
                // Mostrar solo el cuadro seleccionado
                document.getElementById('corte-' + selectedIndex).style.display = 'block';
                printButton.disabled = false;
            } else {
                printButton.disabled = true;
            }
        });
        
        document.getElementById('printButton').addEventListener('click', function() {
            // Activa la función nativa del navegador para Imprimir/Guardar como PDF
            window.print();
        });
    </script>
</body>
</html>