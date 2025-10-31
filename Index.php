<?php
// Incluimos la librería PHPSpreadsheet para leer la hoja Config
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$tramite_options = [];

try {
    $excelFilePath = 'Motos.xlsx';
    $spreadsheet = IOFactory::load($excelFilePath);
    $configWorksheet = $spreadsheet->getSheetByName('Config');

    if ($configWorksheet) {
        // Leemos desde la fila 12 de la Columna D
        $startRow = 12;
        $highestRow = $configWorksheet->getHighestRow();

        for ($row = $startRow; $row <= $highestRow; $row++) {
            $tramite_name = $configWorksheet->getCell('D' . $row)->getValue();
            
            // Validamos que haya valor en D y valores numéricos en E, F, G, H, I, J
            $isValid = true;
            if (empty($tramite_name)) {
                $isValid = false;
            } else {
                // Revisamos si E,F,G,H,I,J tienen valores numéricos (incluyendo 0)
                $columns = ['E', 'F', 'G', 'H', 'I', 'J', 'K'];
                foreach ($columns as $col) {
                    $cellValue = $configWorksheet->getCell($col . $row)->getCalculatedValue();
                    // Usamos is_numeric para incluir el 0. No usamos empty() para no excluir el 0
                    if (!is_numeric($cellValue)) {
                        $isValid = false;
                        break;
                    }
                }
            }
            
            if ($isValid) {
                // Si es válido, guardamos el texto de la opción D y la fila para usarla como ID/valor
                $tramite_options[$row] = $tramite_name;
            }
        }
    }
} catch (\Exception $e) {
    // Si hay un error, el dropdown queda vacío.
    error_log("Error al leer la hoja Config para trámites: " . $e->getMessage());
}

// Generar las opciones HTML para el dropdown dinámico
$tramite_options_html = '<option value="">Seleccione el Trámite...</option>';
foreach ($tramite_options as $row_id => $name) {
    // Usamos el ID de la fila como valor para poder buscar los montos luego en procesar.php
    $tramite_options_html .= "<option value=\"{$row_id}\">{$name}</option>";
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cálculos de Tránsito</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <div class="container-xl"> <h1 class="text-center mt-5">Inscripción de trámites</h1>
        <div class="text-center mb-4">
            <a href="registros.php" class="btn btn-info">Ver Registros</a>
        </div>
        <form class="mt-4" method="POST" action="procesar.php">
            <div class="row">

                <div class="col-md-4 mb-4">
                    <div class="p-3 border rounded bg-light"> <h3>Cliente</h3>
                        <div class="form-group">
                            <label for="nombres">Nombres</label>
                            <input type="text" class="form-control" id="nombres" name="nombres" required>
                        </div>
                         <div class="form-group">
                            <label for="apellidos">Apellidos</label>
                            <input type="text" class="form-control" id="apellidos" name="apellidos" required>
                        </div>
                        <div class="form-group">
                            <label for="cedula">Cédula</label>
                            <input type="number" class="form-control" id="cedula" name="cedula" required>
                        </div>
                        <div class="form-group">
                            <label for="direccion">Dirección</label>
                            <input type="text" class="form-control" id="direccion" name="direccion" required>
                        </div>
                        <div class="form-group">
                            <label for="telefono">Teléfono</label>
                            <input type="text" class="form-control" id="telefono" name="telefono" required>
                        </div>
                    </div>
                </div>

                <div class="col-md-4 mb-4">
                    <div class="p-3 border rounded bg-light">
                        <h3>Acerca de la Moto</h3>
                        
                        <div class="row">
                            
                            <div class="col-md-6">
                                <div class="form-group">
                                    <label for="placa">Placa</label>
                                    <input type="text" class="form-control" id="placa" name="placa" required>
                                </div>
                                <div class="form-group">
                                    <label for="marca">Marca</label>
                                    <input type="text" class="form-control" id="marca" name="marca" required>
                                </div>
                                <div class="form-group">
                                    <label for="linea">Línea</label>
                                    <input type="text" class="form-control" id="linea" name="linea" required>
                                </div>
                                <div class="form-group">
                                    <label for="no_motor">No. Motor</label>
                                    <input type="text" class="form-control" id="no_motor" name="no_motor" required>
                                </div>
                                <div class="form-group">
                                    <label for="no_chasis">No. Chasis</label>
                                    <input type="text" class="form-control" id="no_chasis" name="no_chasis" required>
                                </div>
                            </div>

                            <div class="col-md-6">
                                <div class="form-group">
                                    <label for="ano">Año (AAAA)</label>
                                    <input type="number" class="form-control" id="ano" name="ano" required>
                                </div>
                                <div class="form-group">
                                    <label for="cilindraje">Cilindraje</label>
                                    <input type="number" class="form-control" id="cilindraje" name="cilindraje" required>
                                    <div id="impuesto" class="mt-2"></div>
                                </div>
                                <div class="form-group">
                                    <label for="BI">Precio Base</label>
                                    <input type="number" class="form-control" id="BI" name="BI">
                                </div>
                                 <div class="form-group">
                                    <label for="no_factura">Numero de factura</label>
                                    <input type="number" class="form-control" id="no_fac" name="no_fac">
                                </div>
                            </div>
                        </div> 
                    </div>
                </div>

                <div class="col-md-4 mb-4">
                    
                    <div id="transito_tramite_fields" class="p-3 border rounded bg-light mb-4"> 
                        <h3>Tránsito del Trámite</h3>
                        <div class="form-group">
                            <label for="COC">¿Contado o Credito?</label>
                            <select class="form-control" id="COC" name="COC" required>
                                <option value="">Seleccione...</option>
                                <option value="Contado">Contado</option>
                                <option value="Credito">Credito</option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label for="transito">Seleccione el Tránsito al que pertenece el trámite</label>
                            <select class="form-control" id="transito" name="transito" required>
                                <option value="">Seleccione...</option>
                                <option value="Villa">Villa del Rosario</option>
                                <option value="Cucuta">Cúcuta</option>
                                <option value="Patios">Los Patios</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="p-3 border rounded bg-light"> <h3>Trámite a realizar</h3>
                        <div class="form-group">
                            <label for="tipo_operacion">Tipo de trámite</label>
                            <select class="form-control" id="tipo_operacion" name="tipo_operacion" required>
                                <option value="Inicial">Inscripción inicial</option>
                                <option value="Otro">Otro</option>
                            </select>
                        </div>

                        <div id="otro_tramite_fields" style="display:none;">
                            <div class="form-group">
                                <label for="tipo_tramite_otro">Tipo de Trámite (Otro)</label>
                                <select class="form-control" id="tipo_tramite_otro" name="tipo_tramite_otro">
                                    <?php echo $tramite_options_html; ?>
                                </select>
                            </div>

                            <div class="form-group">
                                <label for="transito_otro">Tránsito para Trámite (Otro)</label>
                                <select class="form-control" id="transito_otro" name="transito_otro">
                                    <option value="">Seleccione...</option>
                                    <option value="Villa">Villa del Rosario</option>
                                    <option value="Cucuta">Cúcuta</option>
                                    <option value="Patios">Los Patios</option>
                                </select>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="text-center mt-4 mb-5"> 
                <button type="submit" class="btn btn-primary btn-lg">Añadir</button>
            </div>
        </form>
    </div>
    
    <div class="container-xl">
        <div class="row">
            <div class="col-md-4 offset-md-4 mb-4">
                <div class="p-3 border rounded bg-light">
                    <h3>Fecha de Corte</h3>
                    <form method="POST" action="corte_semanal.php">
                        <div class="form-group">
                            <label for="corte_switch">¿Establecer una fecha de corte?</label>
                            <select class="form-control" id="corte_switch" name="corte_switch" required>
                                <option value="">Seleccione...</option>
                                <option value="Si">Sí</option>
                                <option value="No">No</option>
                            </select>
                        </div>
                        <button type="submit" class="btn btn-warning btn-block">Establecer fecha de corte</button>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Lógica de impuesto (Solo visual, el cálculo final es en procesar.php)
        document.getElementById('cilindraje').addEventListener('input', function() {
            const cilindraje = parseInt(this.value);
            const impuestoDiv = document.getElementById('impuesto');
            if (cilindraje >= 126) {
                impuestoDiv.innerHTML = '<span class="text-danger">Aplica para impuesto</span>';
            } else {
                impuestoDiv.innerHTML = '<span class="text-success">No aplica para impuesto</span>';
            }
        });

        // Lógica para mostrar/ocultar campos de "Otro Trámite" y "Tránsito del Trámite"
        document.getElementById('tipo_operacion').addEventListener('change', function() {
            const isOtro = this.value === 'Otro';
            
            const transitoTramiteFields = document.getElementById('transito_tramite_fields');
            const otroTramiteFields = document.getElementById('otro_tramite_fields');
            
            // 1. CONTROL VISUAL: Ocultar/Mostrar la sección "Tránsito del Trámite"
            transitoTramiteFields.style.display = isOtro ? 'none' : 'block';
            
            // Mostrar/Ocultar campos de "Otro Trámite"
            otroTramiteFields.style.display = isOtro ? 'block' : 'none';
            
            // 2. CONTROL DE REQUERIDOS Y DATOS
            
            // Control de campos requeridos para "Otro Trámite"
            // Nota: Se usa setAttribute('required', 'required') para asegurar el comportamiento HTML5
            if (isOtro) {
                document.getElementById('tipo_tramite_otro').setAttribute('required', 'required');
                document.getElementById('transito_otro').setAttribute('required', 'required');
                
                // Limpiar y quitar required de Inicial
                document.getElementById('COC').value = '';
                document.getElementById('transito').value = '';
                document.getElementById('COC').removeAttribute('required');
                document.getElementById('transito').removeAttribute('required');

            } else {
                document.getElementById('COC').setAttribute('required', 'required');
                document.getElementById('transito').setAttribute('required', 'required');
                
                // Limpiar y quitar required de Otro
                document.getElementById('tipo_tramite_otro').value = '';
                document.getElementById('transito_otro').value = '';
                document.getElementById('tipo_tramite_otro').removeAttribute('required');
                document.getElementById('transito_otro').removeAttribute('required');
            }
        });
        
        // Inicializar el estado de los campos de "Otro Trámite" al cargar la página
        document.getElementById('tipo_tramite_otro').removeAttribute('required');
        document.getElementById('transito_otro').removeAttribute('required');
        // La sección de Tránsito Inicial debe estar visible por defecto y sus campos requeridos
        document.getElementById('COC').setAttribute('required', 'required');
        document.getElementById('transito').setAttribute('required', 'required');
    </script>
</body>
</html>