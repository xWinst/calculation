import { useRef, useState, createRef } from 'react';
import { Button } from 'components';
import { getBook, getRows, message } from 'utils';
import s from './Main.module.css';

const FileInput = ({ name, submit, reff }) => {
    return (
        <input
            name={name}
            className="hidden"
            type="file"
            ref={reff}
            accept=".xls, .xlsx"
            onChange={submit}
            onClick={e => (e.target.value = null)}
        />
    );
};

const morionDict = {
    410725: 150057,
    606318: 293357,
    598739: 325383,
    782286: 326735,
    145653: 439144,
    882790: 293357,
    797782: 462003,
    153087: 444072,
    559034: 183526,
    541813: 199847,

    381950: 569928,
    181625: 568320,
    689402: 317929,
};

const drugTitle = ['Товар', 'Наименование товара', 'Назва товару', 'Препарат', 'Номенклатура'];
const countTitle = ['Количество', 'Кількість', 'Кіл-сть', 'Кількість Прихід', 'Количество приход'];
const providerTitle = ['Постачальник', 'Поставщик', "Дистриб'ютор"];
const expenseTitle = ['Кол-во', 'Кількість Розхід', 'Продажи', 'Расход-чеки', 'Розх ксть уп'];
const providerNames = [
    'ТОВ"БаДМ(ОСНОВНИЙ)"',
    'СП "Оптіма-Фарм, ЛТД"(ОСНОВНИЙ)',
    'Оптіма-Фарм',
    'Бадм 1',
    'БаДМ ТОВ',
    'Оптіма-Фарм,ЛТД СП',
    'Оптіма-Фарм, ЛТД СП',
    'БаДМ ТзОВ',
    'БаДМ',
    'Оптіма',
    'Оптіма-Фарм 10 дн',
];

const Main = () => {
    const [drugs, setDrugs] = useState();
    const [discountedDrugs, setDiscountedDrugs] = useState([]);
    const [mainTable, setMainTable] = useState();
    const [tableData, setTableData] = useState({});
    const [fileName, setFileName] = useState('result');
    const [dictionary, setDictionary] = useState();
    const [pharmacyTotal, setPharmacyTotal] = useState('');
    const [xlsNames, setXlsNames] = useState({});

    const filePicker = useRef([1, 2, 3, 4, 5].map(createRef));

    const loadFile = i => {
        filePicker.current[i].current.click();
    };

    const loadResultTable = async e => {
        const book = await getBook(e.target.files);
        if (!book) return;
        setXlsNames(prev => ({ ...prev, RT: e.target.files[0].name, IR: null, PR: null }));

        setMainTable(book);
        try {
            const mainPage = book.worksheets[0];

            const isDecada = book.worksheets.length === 2;

            let headRow, totalRow, discountRow;
            let monthColumn, drugsColumn, discontDrugsColumn;

            const rows = mainPage._rows;

            for (let row = 0; row < rows.length; row++) {
                const cells = rows[row]._cells;
                for (let cell = 0; cell < cells.length; cell++) {
                    if (cells[cell]?.value === 'ПРЕПАРАТ') {
                        if (!drugsColumn) drugsColumn = cells[cell]._column._number;
                        else discontDrugsColumn = isDecada ? -1 : cells[cell]._column._number;
                        headRow = row;
                    }

                    if (cells[cell]?.value === 'ИТОГО') {
                        totalRow = row;
                    }
                }
            }

            for (let col = 1; col <= rows[totalRow]._cells.length; col++) {
                if (rows[totalRow].getCell(col)._value?.result === 0) {
                    monthColumn = col;
                    break;
                }
            }

            const billingMonth = rows[headRow].getCell(monthColumn).value.split(' ')[3];
            const drugs = [];
            const discountedDrugs = [];

            for (let row = headRow + 1; row < totalRow; row++) {
                drugs.push({
                    name: rows[row].getCell(drugsColumn).value.toLowerCase(),
                    count: 0,
                    morion: rows[row].getCell(drugsColumn + 1).value,
                });

                if (
                    !isDecada &&
                    !discountRow &&
                    rows[row].getCell(discontDrugsColumn).value?.toLowerCase() ===
                        billingMonth.toLowerCase()
                ) {
                    discountRow = row;
                }
            }

            if (isDecada) {
                const rows2 = book.worksheets[1]._rows;
                let headRow2,
                    isStart = false,
                    isEnd = false;

                for (let row = 0; row < rows2.length; row++) {
                    const cells = rows2[row]?._cells;
                    if (!cells) continue;
                    for (let cell = 0; cell < cells.length; cell++) {
                        if (
                            cells[cell]?.value?.toString().toLowerCase() ===
                            billingMonth.toLowerCase()
                        ) {
                            headRow2 = row;
                            break;
                        }

                        if (headRow2 && cells[cell]?.value === 'ПРЕПАРАТ') {
                            isStart = true;
                            break;
                        }
                        if (isStart && cells[cell]?.value === 'ПРЕПАРАТ') isEnd = true;

                        if (headRow2) {
                            drugs.forEach(({ name, morion }, i) => {
                                if (cells[cell]?.value === morion) {
                                    discountedDrugs.push({
                                        name,
                                        count: 0,
                                        expense: 0,
                                        balance: 0,
                                    });
                                }
                            });
                        }
                    }
                    if (isEnd) break;
                }
            } else {
                for (let row = discountRow + 1; row < totalRow; row++) {
                    if (rows[row].getCell(discontDrugsColumn).value)
                        discountedDrugs.push({
                            name: rows[row].getCell(discontDrugsColumn).value.toLowerCase(),
                            count: 0,
                            expense: 0,
                            balance: 0,
                        });
                    else break;
                }
            }

            setTableData({
                monthColumn,
                headRow,
                totalRow,
                drugsColumn,
                billingMonth,
                discountRow,
                discontDrugsColumn,
                isDecada,
            });

            setDiscountedDrugs(discountedDrugs);

            setDrugs(drugs);
        } catch (e) {
            console.log('e: ', e);
            setDrugs(null);
            setXlsNames(prev => ({ ...prev, RT: null, IR: null, PR: null }));
            message.error('Розрахункова таблиця не пройшла валідацію', e.message);
        }
    };

    const loadComparsionTable = async e => {
        const rows = await getRows(e.target.files);
        if (!rows) return;
        setXlsNames(prev => ({ ...prev, CT: e.target.files[0].name, IR: null, PR: null }));
        const dictionary = [];
        try {
            for (let row = 0; row < rows.length; row++) {
                if (rows[row][0])
                    dictionary.push({
                        rName: rows[row][0].toLowerCase(),
                        clName: rows[row][1]?.toLowerCase(),
                    });
            }

            setDictionary(dictionary);
        } catch (e) {
            console.log('e: ', e);
            setXlsNames(prev => ({ ...prev, CT: null, IR: null, PR: null }));
            message.error('Таблиця порівнянь не пройшла валідацію', e.message);
        }
    };

    const resetDrugsCount = drugsClone => {
        drugsClone.forEach(drug => {
            drug.count = 0;
        });
    };

    const resetDrugsCountFOP = drugsClone => {
        drugsClone.forEach(drug => {
            drug.fop = 0;
        });
    };

    const resetBalance = drugsClone => {
        drugsClone.forEach(drug => {
            drug.expense = 0;
            drug.balance = 0;
            drug.pharmacyCount = 0;
        });
    };

    const loadIncomeReport = async e => {
        const rows = await getRows(e.target.files);
        if (!rows) return;
        setXlsNames(prev => ({ ...prev, IR: e.target.files[0].name }));

        try {
            let drugsNamesColumn, drugsCountColumn, providerColumn;
            let headRow, isDS, isProvider, needPenetration;

            for (let row = 0; row < rows.length; row++) {
                let isEnd = false;
                for (let i = 0; i < drugTitle.length; i++) {
                    if (rows[row].includes(drugTitle[i])) {
                        drugsNamesColumn = rows[row].indexOf(drugTitle[i]);
                        headRow = row;
                        isEnd = true;
                        break;
                    }
                }
                if (isEnd) break;
            }

            for (let i = 0; i < countTitle.length; i++) {
                if (rows[headRow].includes(countTitle[i])) {
                    drugsCountColumn = rows[headRow].indexOf(countTitle[i]);
                    isDS = countTitle[i] === 'Кількість Прихід';
                    isProvider =
                        countTitle[i] === 'Кількість Прихід' ||
                        countTitle[i] === 'Количество приход';

                    needPenetration = !(countTitle[i] === 'Количество приход');
                    if (tableData.isDecada) needPenetration = false;
                }
            }

            for (let i = 0; i < providerTitle.length; i++) {
                if (rows[headRow].includes(providerTitle[i])) {
                    providerColumn = rows[headRow].indexOf(providerTitle[i]);
                }
            }

            const drugsJson = JSON.stringify(drugs);
            const drugsClone = JSON.parse(drugsJson);
            const discountedDrugsJson = JSON.stringify(discountedDrugs);
            const discountedDrugsClone = JSON.parse(discountedDrugsJson);

            resetDrugsCount(drugsClone);
            resetDrugsCount(discountedDrugsClone);

            rows.forEach(value => {
                if (value[drugsNamesColumn]) {
                    const clientName = value[drugsNamesColumn].toString().toLowerCase();
                    const index = dictionary.findIndex(({ clName }) => clName === clientName);
                    if (index >= 0) {
                        let realName = dictionary[index].rName;

                        discountedDrugs.forEach(({ name, count }, i) => {
                            if (
                                realName === name &&
                                (providerNames.includes(value[providerColumn]) || isProvider)
                            ) {
                                discountedDrugsClone[i].count += value[drugsCountColumn];
                                if (!isDS) realName = null;
                            }
                        });

                        drugsClone.forEach(({ name, count }, i) => {
                            if (
                                realName === name &&
                                (providerNames.includes(value[providerColumn]) || isProvider)
                            ) {
                                drugsClone[i].count += value[drugsCountColumn];
                            }
                        });
                    }
                }
            });

            setDrugs(drugsClone);
            setDiscountedDrugs(discountedDrugsClone);
            setTableData(prev => ({ ...prev, needPenetration }));
        } catch (e) {
            console.log('e: ', e);
            setXlsNames(prev => ({ ...prev, IR: null }));
            message.error('Звіт по приходу не пройшов валідацію', e.message);
        }
    };

    const loadFOPReport = async e => {
        const rows = await getRows(e.target.files);
        if (!rows) return;
        setXlsNames(prev => ({ ...prev, FR: e.target.files[0].name }));
        try {
            let morionColumn, countColumn, providerColumn;
            let headRow;

            for (let row = 0; row < rows.length; row++) {
                if (rows[row].includes('Код Моріон')) {
                    morionColumn = rows[row].indexOf('Код Моріон');
                    headRow = row;
                    break;
                }
            }

            for (let i = 0; i < countTitle.length; i++) {
                if (rows[headRow].includes(countTitle[i])) {
                    countColumn = rows[headRow].indexOf(countTitle[i]);
                }
            }

            for (let i = 0; i < providerTitle.length; i++) {
                if (rows[headRow].includes(providerTitle[i])) {
                    providerColumn = rows[headRow].indexOf(providerTitle[i]);
                }
            }

            const drugsJson = JSON.stringify(drugs);
            const drugsClone = JSON.parse(drugsJson);

            resetDrugsCountFOP(drugsClone);

            rows.forEach(value => {
                if (value[morionColumn]) {
                    let morionCode = value[morionColumn];
                    let index = drugs.findIndex(({ morion }) => morion === morionCode);
                    if (index < 0) {
                        morionCode = morionDict[morionCode];
                        index = drugs.findIndex(({ morion }) => morion === morionCode);
                    }
                    if (index >= 0) {
                        drugsClone.forEach(({ morion }, i) => {
                            if (
                                morion === morionCode &&
                                providerNames.includes(value[providerColumn])
                            ) {
                                drugsClone[i].fop += value[countColumn];
                            }
                        });
                    }
                }
            });

            setDrugs(drugsClone);
        } catch (e) {
            console.log('e: ', e);
            setXlsNames(prev => ({ ...prev, FR: null }));
            message.error('Звіт по ФОПам не пройшов валідацію', e.message);
        }
    };

    const loadPharmacyReport = async e => {
        const rows = await getRows(e.target.files);
        if (!rows) return;
        setXlsNames(prev => ({ ...prev, PR: e.target.files[0].name }));

        try {
            let drugsNamesColumn, drugsCountColumn, balanceColumn; //, pharmacyColumn;
            let headRow, reserve;

            for (let row = 0; row < rows.length; row++) {
                let isEnd = false;
                for (let i = 0; i < drugTitle.length; i++) {
                    if (rows[row].includes(drugTitle[i])) {
                        drugsNamesColumn = rows[row].indexOf(drugTitle[i]);
                        headRow = drugTitle[i] === 'Номенклатура' ? row - 1 : row;
                        if (rows[row][drugsNamesColumn + 4] === 'Total') {
                            headRow = row + 1;
                            reserve = true;
                        }
                        isEnd = true;
                        break;
                    }
                }
                if (isEnd) break;
            }

            for (let i = 0; i < expenseTitle.length; i++) {
                if (rows[headRow].includes(expenseTitle[i])) {
                    drugsCountColumn = rows[headRow].indexOf(expenseTitle[i]);
                    balanceColumn =
                        expenseTitle[i] === 'Продажи' || expenseTitle[i] === 'Розх ксть уп'
                            ? drugsCountColumn + 2
                            : drugsCountColumn + 1;
                    break;
                }
            }

            const discountedDrugsJson = JSON.stringify(discountedDrugs);
            const discountedDrugsClone = JSON.parse(discountedDrugsJson);

            resetBalance(discountedDrugsClone);
            rows.forEach(value => {
                if (value[drugsNamesColumn]) {
                    const clientName = value[drugsNamesColumn].toString().toLowerCase();
                    const index = dictionary.findIndex(({ clName }) => clName === clientName);
                    if (index >= 0) {
                        let realName = dictionary[index].rName;

                        discountedDrugs.forEach(({ name, count }, i) => {
                            if (realName === name) {
                                discountedDrugsClone[i].expense += value[drugsCountColumn] || 0;
                                discountedDrugsClone[i].balance += value[balanceColumn] || 0;
                                if (value[balanceColumn] > 0)
                                    discountedDrugsClone[i].pharmacyCount++;
                            }
                        });
                    }
                }
            });

            setDiscountedDrugs(discountedDrugsClone);
            setTableData(prev => ({ ...prev, reserve }));
        } catch (e) {
            console.log('e: ', e);
            setXlsNames(prev => ({ ...prev, PR: null }));
            message.error('Звіт по аптечний не пройшов валідацію', e.message);
        }
    };

    const saveResult = async () => {
        const result = { ...mainTable };
        Object.setPrototypeOf(result, mainTable.__proto__);
        const mainPage = result.worksheets[0];
        const rows = mainPage._rows;

        for (let row = tableData.headRow + 1; row < tableData.totalRow; row++) {
            rows[row].getCell(tableData.monthColumn).value =
                drugs[row - tableData.headRow - 1].count || null;

            if (drugs[row - tableData.headRow - 1].fop !== undefined) {
                rows[row].getCell(tableData.monthColumn + 2).value =
                    drugs[row - tableData.headRow - 1].fop || null;
            }
        }

        if (!tableData.isDecada) {
            const start = tableData.discountRow + 1;
            const end = discountedDrugs.length + start;

            for (let row = start; row < end; row++) {
                rows[row].getCell(tableData.discontDrugsColumn + 2).value =
                    discountedDrugs[row - start].count;
                rows[row].getCell(tableData.discontDrugsColumn + 5).value =
                    discountedDrugs[row - start].expense;

                if (tableData.needPenetration) {
                    const shift = tableData.reserve ? 2 : 0;

                    const penetration =
                        (100 * discountedDrugs[row - start].pharmacyCount) / pharmacyTotal;
                    const discount =
                        penetration < 60
                            ? 0
                            : penetration < 75
                            ? 0.03
                            : penetration < 90
                            ? 0.04
                            : 0.05;
                    if (tableData.reserve) {
                        const balance = discountedDrugs[row - start].balance;
                        const dis2 =
                            balance / pharmacyTotal < 1
                                ? 0
                                : balance / pharmacyTotal < 3
                                ? 0.02
                                : 0.05;

                        rows[row].getCell(tableData.discontDrugsColumn + 8).value = {
                            formula: `${balance}/${pharmacyTotal}`,
                        };
                        rows[row].getCell(tableData.discontDrugsColumn + 9).value = {
                            formula: `${
                                rows[row].getCell(tableData.discontDrugsColumn + 3).address
                            } * ${dis2}`,
                        };
                    }

                    rows[row].getCell(tableData.discontDrugsColumn + 8 + shift).value = {
                        formula: `${
                            discountedDrugs[row - start].pharmacyCount
                        }/${pharmacyTotal} * 100`,
                    };

                    rows[row].getCell(tableData.discontDrugsColumn + 9 + shift).value = {
                        formula: `${
                            rows[row].getCell(tableData.discontDrugsColumn + 3).address
                        } * ${discount}`,
                    };
                }
            }
        } else {
            const rows2 = result.worksheets[1]._rows;
            let headRow2,
                isStart = false,
                isEnd = false;

            for (let row = 0; row < rows2.length; row++) {
                const cells = rows2[row]?._cells;
                if (!cells) continue;
                for (let cell = 0; cell < cells.length; cell++) {
                    if (
                        cells[cell]?.value?.toString().toLowerCase() ===
                        tableData.billingMonth.toLowerCase()
                    ) {
                        headRow2 = row;
                        break;
                    }

                    if (headRow2 && cells[cell]?.value === 'ПРЕПАРАТ') {
                        isStart = true;
                        break;
                    }
                    if (isStart && cells[cell]?.value === 'ПРЕПАРАТ') isEnd = true;

                    if (headRow2) {
                        drugs.forEach(({ name, morion }, i) => {
                            if (cells[cell]?.value === morion) {
                                const drug = discountedDrugs.find(drug => drug.name === name);

                                cells[cell + 2].value = drug.count;
                            }
                        });
                    }
                }
                if (isEnd) break;
            }
        }

        result.calcProperties.fullCalcOnLoad = true;

        const buffer = await result.xlsx.writeBuffer();

        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${fileName}.xlsx`;
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    };

    const setTotal = e => {
        let { value } = e.target;
        if (!Number(value) && value) return;
        if (value < 1) value = '';
        setPharmacyTotal(value);
    };

    return (
        <div className={s.container}>
            <div className={s.flexBox}>
                <p className={s.p} style={{ color: xlsNames.RT ? '#00bd3b' : '#ff0000' }}>
                    {xlsNames.RT || 'Не завантажено'}
                </p>
                <Button click={() => loadFile(0)} text="Завантажити розрахункову таблицю" />
                {xlsNames.RT && (
                    <>
                        <p>Розрахунковий місяць: </p>
                        <p>{tableData.billingMonth}</p>
                    </>
                )}
            </div>
            <div className={s.flexBox}>
                <p className={s.p} style={{ color: xlsNames.CT ? '#00bd3b' : '#ff0000' }}>
                    {xlsNames.CT || 'Не завантажено'}
                </p>

                <Button click={() => loadFile(3)} text="Завантажити таблицю порівняння" />
            </div>
            <div className={s.flexBox}>
                <p className={s.p} style={{ color: xlsNames.IR ? '#00bd3b' : '#ff0000' }}>
                    {xlsNames.IR || 'Не завантажено'}
                </p>
                <Button
                    click={() => loadFile(1)}
                    text="Завантажити звіт по приходах"
                    dis={!xlsNames.RT || !xlsNames.CT}
                />
            </div>
            <div className={s.flexBox}>
                <p className={s.p} style={{ color: xlsNames.PR ? '#00bd3b' : '#ff0000' }}>
                    {xlsNames.PR || 'Не завантажено'}
                </p>
                <Button
                    click={() => loadFile(2)}
                    text="Завантажити звіт по aптечний"
                    dis={!xlsNames.RT || !xlsNames.CT}
                />
                <p>Загальна кількість аптек -</p>
                <input className={s.input} onChange={setTotal} value={pharmacyTotal} />
            </div>
            <div className={s.flexBox}>
                <p className={s.p} style={{ color: xlsNames.FR ? '#00bd3b' : '#ff0000' }}>
                    {xlsNames.FR || 'Не завантажено'}
                </p>
                <Button
                    click={() => loadFile(4)}
                    text="Завантажити звіт по ФОПам"
                    dis={!xlsNames.RT || !xlsNames.CT}
                />
            </div>
            <FileInput name="RT" submit={loadResultTable} reff={filePicker.current[0]} />
            <FileInput name="IR" submit={loadIncomeReport} reff={filePicker.current[1]} />
            <FileInput name="PR" submit={loadPharmacyReport} reff={filePicker.current[2]} />
            <FileInput name="CT" submit={loadComparsionTable} reff={filePicker.current[3]} />
            <FileInput name="FR" submit={loadFOPReport} reff={filePicker.current[4]} />
            <div className={s.flexBox + ' ' + s.mt}>
                <p className={s.end}>Ім'я файлу:</p>
                <input onChange={e => setFileName(e.target.value)} value={fileName} />
                <Button
                    click={saveResult}
                    text="Отримати результат"
                    dis={
                        !xlsNames.RT ||
                        !xlsNames.CT ||
                        !xlsNames.IR ||
                        (!xlsNames.PR && !tableData.isDecada) ||
                        (!pharmacyTotal && tableData.needPenetration)
                    }
                />
            </div>
        </div>
    );
};

export default Main;
