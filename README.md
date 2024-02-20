# calc-file-object-converter

## Archivos soportados
- xls
- xlsx
 
## Data types soportados
- String
- Integer
- int
- Double
- double
- Long
- long
- Boolean
- boolean
- LocalDate 

## Ejemplo de uso
En el directorio base se encuentran dos archivos utilizados para
las pruebas al igual que sus dos correspondientes clases de Java.

``` java
// La variable file es el MultipartFile recibido en el controlador
Workbook workbook = CalcFileConverter.getWorkbook(file);
Sheet sheet;

// Sheet es cada hoja del archivo
// Se puede obtener una Sheet en base a su nombre o 
// en base a su indice dentro del archivo (empieza en 0)
if(sheetName != null && !sheetName.equalsIgnoreCase("")){
    sheet = CalcFileConverter.getSheetByName(workbook, sheetName);
}else{
    sheet = CalcFileConverter.getSheetByIndex(workbook, sheetIndex);
}

if(sheet == null) throw new SheetNotFoundException();

// Para obtener una tabla solo hace falta obtener su esquina
// superior-izquierda, la cual conforma el encabezado de la misma
// y saber el largo y alta de la tabla, pero para eso hay una función
// que permite conocer estos datos corroborando donde se encuentra el primer 
// espacio en blanco (fin de la tabla) 
int endRow = CalcFileConverter.getLastRow(sheet, startRow, startColumn);
int endColumn = CalcFileConverter.getLastColumn(sheet, startRow, startColumn);

// Este ultimo metodo extrae la lista de objetos genericos convertidos del archivo
List<TestUserModel> users = CalcFileConverter.extractObjectsFromTable(
        sheet, 
        startRow, 
        endRow, 
        startColumn, 
        endColumn, 
        TestUserModel.class
);
```
