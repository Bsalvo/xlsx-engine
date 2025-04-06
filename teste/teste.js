const Excel = require('../src/Excel');
const E = new Excel('C:/Users/bruno/Downloads');

async function teste(){

    await E.create('teste', ['Nome'],  [{nome: 'Bruno'}]);

}

teste();