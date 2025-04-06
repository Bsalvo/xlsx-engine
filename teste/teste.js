const Excel = require('../src/Excel');
const E = new Excel('C:/Users/bruno/Downloads');

async function teste(){

    let teste = await E.toJson('matriz_15150189.xlsx', '15150189', 2, 3);
    console.log(teste.data);

}

teste();