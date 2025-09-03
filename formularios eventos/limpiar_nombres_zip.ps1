
# Script para renombrar archivos y carpetas eliminando caracteres problemáticos para compresión ZIP
$rootPath = Get-Location

# Caracteres problemáticos a eliminar o reemplazar
$replacements = @{
    '–' = '-'    # Guion largo a guion normal
    '’' = "'"    # Comilla curva a comilla simple
    '“' = '"'    # Comillas dobles curvas
    '”' = '"'
    '#' = ''     # Eliminar #
    '¿' = ''
    '¡' = ''
    '`' = ''
    ':' = ''
    ';' = ''
    ',' = ''
    '\\' = ''
    '/' = ''
    '\|' = ''
    '\*' = ''
    '\?' = ''
    '<' = ''
    '>' = ''
    '\"' = ''
}

# Función para limpiar nombres
function Clean-Name($name) {
    $cleaned = $name
    foreach ($key in $replacements.Keys) {
        $cleaned = $cleaned -replace $key, $replacements[$key]
    }
    return $cleaned
}

# Renombrar archivos
Get-ChildItem -Recurse -File | ForEach-Object {
    $newName = Clean-Name $_.Name
    if ($_.Name -ne $newName) {
        Write-Host "Renombrando archivo:" $_.Name "->" $newName
        Rename-Item -LiteralPath $_.FullName -NewName $newName
    }
}

# Renombrar carpetas (de más profundas a más superficiales)
Get-ChildItem -Recurse -Directory | Sort-Object FullName -Descending | ForEach-Object {
    $newName = Clean-Name $_.Name
    if ($_.Name -ne $newName) {
        Write-Host "Renombrando carpeta:" $_.Name "->" $newName
        Rename-Item -LiteralPath $_.FullName -NewName $newName
    }
}
