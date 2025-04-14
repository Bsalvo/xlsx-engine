const Excel = require('../src/Excel');
const E = new Excel('C:/Users/bruno/Downloads');

async function teste() {

    await E.create(
        'Exemplo',
        [
            'Nome',
            {
                value: 'Status',
                key: 'status',
                validation: {
                    select: true,
                    values: ['Ativo', 'Inativo', 'Pendente'],
                },
            },
            {
                value: 'Coragem',
                key: 'coragem',
                validation: {
                    select: true,
                    values: ['Cão', 'Covarde'],
                },
            },
        ],
        [
            { nome: 'Bruno', status: 'Ativo' },
            { nome: 'Maria', coragem: 'Covarde' },
        ],
    );
}

teste();