# calc-file-object-converter

## Archivos soportados
- xls
- xlsx
- xlsm
 
## Tipos de datos soportados por defecto
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

## Añadiendo soporte a otros tipos de datos
### Dato literal en Excel
```
contactless_modes = Contactless Magstripe (MS Mode)
```
### Clase representada en Java
``` java
public class CriteriaModel {

    private String name;
    private String value;

    public CriteriaModel(String name, String value) {
        this.name = name;
        this.value = value;
    }
}
```
### Registrar la conversion de datos
``` java
CalcFileConverter.addDataTypeConversion(CriteriaModel.class, data -> {
    String[] dataSplit = data.split(" = ", 2);

    if(dataSplit.length < 2) {
        return new CriteriaModel(null, data);
    }else{
        String criteriaName = dataSplit[0];
        String criteriaValue = dataSplit[1];

        return new CriteriaModel(criteriaName, criteriaValue);
    }
});
```

## Anotación
``` java
@ExcelColumn("Nombre de la columna en la tabla del Excel")
```

## Ejemplo de uso
En el directorio base se encuentran dos archivos utilizados para
las pruebas al igual que sus dos correspondientes clases de Java.
### Clase representada en Java sobre cada Row del Excel
``` java
public class TestUserModel {
    @ExcelColumn("ID")
    private int id;
    @ExcelColumn("Name")
    private String name;
    @ExcelColumn("Email")
    private String email;
}
```
### Código del controlador
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
int startRow = 0;
int startColumn = 0; 
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