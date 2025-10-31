<?php
require 'vendor/autoload.php'; // Incluye el autoloader de Composer

use PhpOffice\PhpSpreadsheet\IOFactory;

// --- CONFIGURACIÓN DE EXCEL ---
$excelFilePath = 'Motos.xlsx'; 
$sheetName = 'Inscripciones';
$startRow = 9; // Fila donde inician los datos (después de encabezados)

// Lista de todas las columnas de A a X
$COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X'];
$HEADERS = [
    'A' => 'No. Trámite', 'B' => 'Nombres', 'C' => 'Apellidos', 'D' => 'Cedula', 'E' => 'Dirección', 'F' => 'Teléfono',
    'G' => 'No. Factura', 'H' => 'Placa', 'I' => 'Marca', 'J' => 'Línea', 'K' => 'No. Motor', 'L' => 'No. Chasis',
    'M' => 'Año', 'N' => 'Cilindraje', 'O' => 'Base Impuesto', 'P' => 'Impuesto', 'Q' => 'Tipo Pago', 'R' => 'Tránsito',
    'S' => 'Caja', 'T' => 'CUPL', 'U' => 'Fecha', 'V' => 'Comisión', 'W' => 'Razón Trámite', 'X' => 'Total'
];

$data_all_rows = []; 
$registros_filtrables = []; 
$marcas = [];
$fechas = [];
$comisiones = [];
$razones = [];
$semanas_corte = []; 
$corte_rows_data = []; 
$total_registros_filtrables = 0; 

try {
    // 1. Cargar el archivo de Excel
    $spreadsheet = IOFactory::load($excelFilePath);
    $worksheet = $spreadsheet->getSheetByName($sheetName);

    if (!$worksheet) {
        throw new Exception("La hoja de cálculo '{$sheetName}' no fue encontrada.");
    }

    $highestRow = $worksheet->getHighestRow();

    // 2. Recorrer todos los datos y recolectar valores para filtros
    for ($row_num = $startRow; $row_num <= $highestRow; $row_num++) {
        $cellA = $worksheet->getCell('A' . $row_num)->getCalculatedValue();
        $is_corte = (strtoupper($cellA) === 'CORTE');

        // CORRECCIÓN: Omitir la fila si la columna A (ID) está completamente vacía y no es un corte
        if (empty($cellA) && !$is_corte) {
            continue; 
        }

        // Almacenar todos los datos de la fila (de A a X)
        $registro = ['is_corte' => $is_corte, 'row_num' => $row_num]; 
        foreach ($COLUMNS as $col) {
            $registro[$col] = $worksheet->getCell($col . $row_num)->getCalculatedValue();
        }
        
        $data_all_rows[] = $registro;

        // Recolección de filtros (solo transacciones)
        if (!$is_corte && !empty($registro['A'])) {
            $total_registros_filtrables++;
            
            $marca = strtoupper($registro['I']);
            $fecha = $worksheet->getCell('U' . $row_num)->getFormattedValue(); 
            $comision = (string)(float)$registro['V'];
            $razon = $registro['W'];

            $marcas[$marca] = true;
            $fechas[$fecha] = true;
            $comisiones[$comision] = true;
            $razones[$razon] = true;
            
            $registros_filtrables[] = $registro;

        } elseif ($is_corte) {
            // Recolección de Semanas de Corte (Columna U)
            $semana_nombre = $worksheet->getCell('U' . $row_num)->getCalculatedValue();
            if (!empty($semana_nombre)) {
                $semanas_corte[$semana_nombre] = true;
                $corte_rows_data[] = ['row_num' => $row_num, 'nombre' => $semana_nombre];
            }
        }
    }

    // Preparar arrays para el HTML de filtros
    $marcas = array_keys($marcas); sort($marcas);
    $fechas = array_keys($fechas); sort($fechas);
    $comisiones = array_keys($comisiones); sort($comisiones);
    $razones = array_keys($razones); sort($razones);
    $semanas_corte = array_keys($semanas_corte); 

} catch (Exception $e) {
    $error_message = "Error al cargar o leer el archivo Excel: " . $e->getMessage();
    $data_all_rows = []; 
}

// 3. Aplicar Filtros (si se enviaron por GET)
$data_final_a_mostrar = [];

// Chequear si se aplicó AL MENOS un filtro
$filtros_activos = isset($_GET['marca']) || isset($_GET['impuesto']) || isset($_GET['fecha']) || isset($_GET['comision']) || isset($_GET['razon']) || isset($_GET['semana_corte']) || isset($_GET['no_factura']); // MODIFICADO
$semana_corte_filtro = $_GET['semana_corte'] ?? '';

// --- INICIO DEL CONDICIONAL AÑADIDO ---
// Verificar si hay filtros activos que NO sean el de semana_corte (para ocultar los cortes).
$otros_filtros_activos = (
    !empty($_GET['marca']) || 
    !empty($_GET['impuesto']) || 
    !empty($_GET['fecha']) || 
    !empty($_GET['comision']) || 
    !empty($_GET['razon']) ||
    !empty($_GET['no_factura']) // MODIFICADO
);
// --- FIN DEL CONDICIONAL AÑADIDO ---

if (!$filtros_activos) {
    $data_final_a_mostrar = $data_all_rows;
} else {
    $marca_filtro = strtoupper($_GET['marca'] ?? '');
    $impuesto_filtro = $_GET['impuesto'] ?? '';
    $fecha_filtro = $_GET['fecha'] ?? '';
    $comision_filtro = $_GET['comision'] ?? '';
    $razon_filtro = $_GET['razon'] ?? '';
    $no_factura_filtro = $_GET['no_factura'] ?? ''; // NUEVA VARIABLE

    // Lógica para filtrar por Semana de Corte
    $row_start = $startRow; // Fila de inicio por defecto
    $row_end = $highestRow; // Fila final por defecto
    $corte_a_mostrar = null; // Fila de corte específica para mostrar

    if (!empty($semana_corte_filtro)) {
        $corte_index = -1;
        // 1. Encontrar el índice del corte seleccionado
        foreach ($corte_rows_data as $index => $corte) {
            if ($corte['nombre'] === $semana_corte_filtro) {
                $corte_index = $index;
                $corte_a_mostrar = $corte['row_num'];
                break;
            }
        }
        
        if ($corte_index !== -1) {
            // 2. Establecer el límite superior (la fila de corte seleccionada)
            $row_end = $corte_rows_data[$corte_index]['row_num'] - 1;

            // 3. Establecer el límite inferior (la fila después del corte anterior)
            if ($corte_index > 0) {
                $row_start = $corte_rows_data[$corte_index - 1]['row_num'] + 1;
            } else {
                // Si es el primer corte, el inicio es $startRow
                $row_start = $startRow;
            }
        }
    }


    // 4. Filtrar los registros de transacciones aplicando todos los demás filtros
    $registros_filtrados = array_filter($registros_filtrables, function($registro) use ($marca_filtro, $impuesto_filtro, $fecha_filtro, $comision_filtro, $razon_filtro, $no_factura_filtro, $worksheet, $row_start, $row_end) { // MODIFICADO: Agregada $no_factura_filtro
        
        // Aplicar filtro de rango de filas si se seleccionó una semana de corte
        if ($registro['row_num'] < $row_start || $registro['row_num'] > $row_end) {
            return false;
        }

        // Columna I (Marca)
        if ($marca_filtro !== '' && strtoupper($registro['I']) !== $marca_filtro) {
            return false;
        }

        // Columna P (Impuesto)
        $impuesto_val = (float)$registro['P'];
        if ($impuesto_filtro === 'con_impuesto' && $impuesto_val <= 0) {
            return false;
        }
        if ($impuesto_filtro === 'sin_impuesto' && $impuesto_val > 0) {
            return false;
        }

        // Columna U (Fecha)
        try {
            $fecha_registro = $worksheet->getCell('U' . $registro['row_num'])->getFormattedValue();
        } catch (\Exception $e) {
            $fecha_registro = 'ERROR';
        }

        if ($fecha_filtro !== '' && $fecha_registro !== $fecha_filtro) {
            return false;
        }
        
        // Columna V (Comisión)
        $comision_val = (string)(float)$registro['V'];
        if ($comision_filtro !== '' && $comision_val !== $comision_filtro) {
            return false;
        }

        // Columna W (Razón de Trámite)
        if ($razon_filtro !== '' && $registro['W'] !== $razon_filtro) {
            return false;
        }

        // NUEVO FILTRO: Columna G (No. Factura)
        if ($no_factura_filtro !== '' && (string)$registro['G'] !== $no_factura_filtro) {
            return false;
        }

        return true;
    });

    // 5. Reconstruir la lista final para mostrar
    $filtered_row_nums = array_column($registros_filtrados, 'row_num');
    
    foreach ($data_all_rows as $row) {
        if ($row['is_corte']) {
            
            $mostrar_corte = false;
        
            if (!empty($semana_corte_filtro)) {
                // Caso 1: Se seleccionó una semana de corte específica.
                // Siempre mostramos la fila de corte seleccionada.
                if ($row['row_num'] === $corte_a_mostrar) {
                    $mostrar_corte = true;
                }
            } else {
                // Caso 2: NO se seleccionó una semana de corte.
                // Solo mostramos todos los cortes si NO hay NINGÚN otro filtro activo.
                if (!$otros_filtros_activos) {
                    $mostrar_corte = true;
                }
            }

            if ($mostrar_corte) {
                $data_final_a_mostrar[] = $row;
            }
            
        } else {
            // Incluir solo las filas de datos filtradas
            if (in_array($row['row_num'], $filtered_row_nums)) {
                $data_final_a_mostrar[] = $row;
            }
        }
    }
}

// Contar registros de transacciones mostrados para el encabezado
$count_mostrar = 0;
foreach($data_final_a_mostrar as $row) {
    if (!$row['is_corte'] && !empty($row['A'])) {
        $count_mostrar++;
    }
}
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📋 Historial de Registros Completo</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="styles.css">
    <style>
        .container-fluid { padding: 20px; }
        .filter-form { margin-bottom: 20px; padding: 15px; border: 1px solid #dee2e6; border-radius: 5px; background-color: #f8f9fa; }
        .table-responsive { max-height: 80vh; overflow-y: auto; } 
        th, td { white-space: nowrap; font-size: 11px; padding: 5px !important; } 
        
        /* Fondo azul claro para las filas de transacción */
        /* Aplicado a 'tr' en lugar de 'tbody tr' para que 'fila-transaccion' funcione */
        .table-striped tr.fila-transaccion td {
            background-color: #e6f7ff !important;
        }
        .table-striped tr.fila-transaccion:nth-of-type(odd) td {
            background-color: #cceeff !important;
        }

        /* Estilo de fila de CORTE (mantenido) */
        .corte-row td { 
            background-color: #ffeeba !important; /* !important para sobreescribir table-striped */
            font-weight: bold; 
            text-align: right; 
            border-top: 2px solid #ffc107;
            border-bottom: 2px solid #ffc107;
        }
        .corte-row td:first-child { 
             text-align: center; 
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <h1 class="text-center">📋 Historial de Registros</h1>
        
        <?php if (isset($error_message)): ?>
            <div class="alert alert-danger" role="alert"><?= $error_message ?></div>
        <?php endif; ?>

        <div class="filter-form">
            <form method="GET">
                <div class="form-row">
                    <div class="col-md-2 form-group">
                        <label for="marca">Filtro por Marca</label>
                        <select class="form-control" id="marca" name="marca">
                            <option value="">Todas las Marcas</option>
                            <?php foreach ($marcas as $m): ?>
                                <option value="<?= $m ?>" <?= (isset($_GET['marca']) && strtoupper($_GET['marca']) === $m) ? 'selected' : '' ?>>
                                    <?= $m ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    
                    <div class="col-md-2 form-group">
                        <label for="impuesto">Filtro por Impuesto</label>
                        <select class="form-control" id="impuesto" name="impuesto">
                            <option value="">Todos</option>
                            <option value="con_impuesto" <?= (isset($_GET['impuesto']) && $_GET['impuesto'] === 'con_impuesto') ? 'selected' : '' ?>>Con Impuesto ( > 0)</option>
                            <option value="sin_impuesto" <?= (isset($_GET['impuesto']) && $_GET['impuesto'] === 'sin_impuesto') ? 'selected' : '' ?>>Sin Impuesto ( = 0)</option>
                        </select>
                    </div>

                    <div class="col-md-2 form-group">
                        <label for="fecha">Filtro por Fecha</label>
                        <select class="form-control" id="fecha" name="fecha">
                            <option value="">Todas las Fechas</option>
                            <?php foreach ($fechas as $f): ?>
                                <option value="<?= $f ?>" <?= (isset($_GET['fecha']) && $_GET['fecha'] === $f) ? 'selected' : '' ?>>
                                    <?= $f ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    
                    <div class="col-md-2 form-group">
                        <label for="semana_corte">Filtro por Semana de Corte</label>
                        <select class="form-control" id="semana_corte" name="semana_corte">
                            <option value="">Todos los Registros</option>
                            <?php foreach ($semanas_corte as $s): ?>
                                <option value="<?= $s ?>" <?= (isset($_GET['semana_corte']) && $_GET['semana_corte'] === $s) ? 'selected' : '' ?>>
                                    <?= $s ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    
                    <div class="col-md-2 form-group">
                        <label for="comision">Filtro por Comisión</label>
                        <select class="form-control" id="comision" name="comision">
                            <option value="">Todas las Comisiones</option>
                            <?php foreach ($comisiones as $c): ?>
                                <option value="<?= $c ?>" <?= (isset($_GET['comision']) && $_GET['comision'] === $c) ? 'selected' : '' ?>>
                                    $<?= number_format((float)$c, 2) ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>

                    <div class="col-md-2 form-group">
                        <label for="razon">Filtro por Razón Trámite</label>
                        <select class="form-control" id="razon" name="razon">
                            <option value="">Todas las Razones</option>
                            <?php foreach ($razones as $r): ?>
                                <option value="<?= $r ?>" <?= (isset($_GET['razon']) && $_GET['razon'] === $r) ? 'selected' : '' ?>>
                                    <?= $r ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                </div> 
                
                <div class="form-row mt-1 justify-content-center">
                    <div class="col-md-2 form-group">
                        <label for="no_factura">Filtro por No. Factura</label>
                        <input type="text" class="form-control" id="no_factura" name="no_factura" 
                               value="<?= htmlspecialchars($_GET['no_factura'] ?? '') ?>" placeholder="Escriba el número">
                    </div>
                </div>
                <div class="form-row mt-3 justify-content-center">
                    <div class="col-md-4 d-flex justify-content-center">
                        <button type="submit" class="btn btn-primary mr-2">Aplicar Filtros</button>
                        <a href="registros.php" class="btn btn-secondary">Limpiar Filtros</a>
                    </div>
                </div>
            </form>
        </div>
        <p class="text-center">
            <a href="cuenta_cobro.php" class="btn btn-warning mr-2">Generar Cuenta de Cobro</a>
        </p>
        <h2>Resultados: <?= $count_mostrar ?> registro(s) de Trámites</h2>
        <div class="table-responsive">
            <table class="table table-striped table-sm table-bordered">
                <thead class="thead-dark" style="position: sticky; top: 0; z-index: 1;">
                    <tr>
                        <?php foreach ($HEADERS as $col => $header): ?>
                            <th><?= $header . " ({$col})" ?></th>
                        <?php endforeach; ?>
                    </tr>
                </thead>
                <tbody>
                    <?php if (empty($data_final_a_mostrar)): ?>
                        <tr><td colspan="<?= count($COLUMNS) ?>" class="text-center">No hay registros que coincidan con los filtros aplicados.</td></tr>
                    <?php else: ?>
                        <?php foreach ($data_final_a_mostrar as $reg): ?>
                            <?php if ($reg['is_corte']): ?>
                                <tr class="corte-row">
                                    <?php foreach ($COLUMNS as $col): ?>
                                        <?php 
                                            $value = $reg[$col];
                                            // Formatear valores numéricos de las columnas de sumatoria, EXCLUYENDO 'O' (Base Impuesto)
                                            $cols_sumatoria = ['P', 'S', 'T', 'V', 'X']; // Impuesto, Caja, CUPL, Comisión, Total
                                            if (in_array($col, $cols_sumatoria)) { 
                                                $value = '$' . number_format((float)$value, 2);
                                            }
                                        ?>
                                        <td><?= htmlspecialchars($value) ?></td>
                                    <?php endforeach; ?>
                                </tr>
                            <?php else: ?>
                                <tr class="fila-transaccion"> 
                                    <?php foreach ($COLUMNS as $col): ?>
                                        <?php 
                                            $value = $reg[$col];
                                            // Formatear valores numéricos para registros de transacción
                                            if (in_array($col, ['O', 'P', 'S', 'T', 'V', 'X'])) { // Moneda/Valores (Incluimos 'O' aquí)
                                                $value = '$' . number_format((float)$value, 2);
                                            } elseif ($col === 'U') { // Fecha
                                                // Obtener el valor formateado de la fecha directamente del objeto $worksheet
                                                try {
                                                    $value = $worksheet->getCell($col . $reg['row_num'])->getFormattedValue();
                                                } catch (\Exception $e) {
                                                    // En caso de error, mostrar el valor sin formato
                                                }
                                            }
                                        ?>
                                        <td><?= htmlspecialchars($value) ?></td>
                                    <?php endforeach; ?>
                                </tr>
                            <?php endif; ?>
                        <?php endforeach; ?>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
        </div>
        <p class="text-center">
            <a href="Index.php" class="btn btn-secondary">Volver al Formulario</a>
        </p>
</body>
</html>