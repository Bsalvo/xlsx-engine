const Excel = require('../src/Excel');
const E = new Excel('/Downloads', { protection: { enabled: true, password: 'segredo123' } });

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
                editable: true,
            },
            {
                value: 'Coragem',
                key: 'coragem',
                validation: {
                    select: true,
                    values: ['Cão', 'Covarde'],
                },
                editable: true,
            },
        ],
        [
            { nome: 'Bruno' },
            { nome: 'Maria', coragem: 'Covarde' },
        ],
    );
}

teste();