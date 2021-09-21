const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

const data = [
    {
       "name":"Teste",
       "email":"teste@gmail.com",
       "cellphone":"1234567890"
    },
    {
       "name":"Pessoa 2",
       "email":"pessoa@gmail.com",
       "cellphone":"1234567899"
    }
   ];

   const headingColumnNames = [
    "Nome",
    "Email",
    "Celular",
]

let headingColumnIndex = 1; //diz que começará na primeira linha
headingColumnNames.forEach(heading => { //passa por todos itens do array
    // cria uma célula do tipo string para cada título
    ws.cell(1, headingColumnIndex++).string(heading);
});


let rowIndex = 2; //começa na linha 2
data.forEach(record => { //passa por cada item do data
    let columnIndex = 1; //diz para começar na primeira coluna
    //transforma cada objeto em um array onde cada posição contém as chaves do objeto (name, email, cellphone)
    Object.keys(record).forEach(columnName =>{
        //cria uma coluna do tipo string para cada item
        ws.cell(rowIndex,columnIndex++)
            .string(record [columnName])
    });
    rowIndex++; //incrementa o contador para ir para a próxima linha
});

wb.write('ArquivoExcel.xlsx');