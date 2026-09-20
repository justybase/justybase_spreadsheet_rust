/**
 * Reverse parity: the TypeScript readers (plus exceljs as an independent
 * third-party implementation) read files produced by the Rust writers.
 *
 * Usage: ts-node ... read_rust_files.ts <rustFixturesDir>
 * Exit 0 = all assertions hold; prints PASS/FAIL lines for the report.
 */
import * as path from 'path';
import { XlsbReader } from '../../justybase_spreadsheet_tasks/src/XlsbReader';
import { XlsxReader } from '../../justybase_spreadsheet_tasks/src/XlsxReader';
// eslint-disable-next-line @typescript-eslint/no-require-imports
const ExcelJS = require('exceljs');

const dir = process.argv[2] ?? '/tmp/rs_fix';
let failures = 0;

function check(name: string, cond: boolean, detail?: string): void {
    if (cond) {
        console.log(`PASS ${name}`);
    } else {
        failures++;
        console.log(`FAIL ${name}${detail ? ` :: ${detail}` : ''}`);
    }
}

function ms(d: unknown): number | null {
    return d instanceof Date ? d.getTime() : null;
}

async function drainXlsb(file: string): Promise<{ names: string[]; rows: any[][] }> {
    const r = new XlsbReader();
    await r.open(file);
    const names = r.getSheetNames();
    const rows: any[][] = [];
    while (await r.read()) {
        const row: any[] = [];
        for (let i = 0; i < r.fieldCount; i++) row.push(r.getValue(i));
        rows.push(row);
    }
    return { names, rows };
}

async function drainXlsx(file: string): Promise<{ names: string[]; rows: any[][] }> {
    const r = new XlsxReader();
    await r.open(file);
    const names = r.getSheetNames();
    const rows: any[][] = [];
    while (await r.read()) {
        const row: any[] = [];
        for (let i = 0; i < r.fieldCount; i++) row.push(r.getValue(i));
        rows.push(row);
    }
    await r.close();
    return { names, rows };
}

async function main(): Promise<void> {
    // --- TS readers over Rust files ---
    const b = await drainXlsb(path.join(dir, 'rs-basic.xlsb'));
    check('xlsb sheet names', JSON.stringify(b.names) === JSON.stringify(['Sheet1']), b.names.join(','));
    check('xlsb row count', b.rows.length === 4, String(b.rows.length));
    check('xlsb header', JSON.stringify(b.rows[0]) === JSON.stringify(['ID', 'Name', 'Count', 'Score', 'Active', 'Joined', 'Amount', 'Big', 'Missing']));
    check('xlsb int', b.rows[1][0] === 1 && b.rows[1][2] === 42, JSON.stringify(b.rows[1].slice(0, 4)));
    check('xlsb unicode', b.rows[1][1] === 'Alicja żółć', String(b.rows[1][1]));
    check('xlsb float/bool', b.rows[1][3] === 12.5 && b.rows[1][4] === true);
    check('xlsb date', ms(b.rows[1][5]) === Date.UTC(2024, 0, 15, 12, 30, 45), String(ms(b.rows[1][5])));
    check('xlsb formatted number', b.rows[1][6] === 1234.5, String(b.rows[1][6]));
    check('xlsb bigint-as-string', b.rows[1][7] === '9007199254740993', String(b.rows[1][7]));
    check('xlsb xml-escapes', b.rows[2][1] === 'Bob & <Co> "q" \'x\'', String(b.rows[2][1]));
    check('xlsb beyond-RK int', b.rows[2][2] === 536870912, String(b.rows[2][2]));
    check('xlsb NaN double', typeof b.rows[3][3] === 'number' && Number.isNaN(b.rows[3][3]), String(b.rows[3][3]));
    check('xlsb epoch date', ms(b.rows[3][5]) === Date.UTC(1899, 11, 30), String(ms(b.rows[3][5])));
    check('xlsb RK lower bound', b.rows[3][2] === -(1 << 29), String(b.rows[3][2]));

    const x = await drainXlsx(path.join(dir, 'rs-basic.xlsx'));
    check('xlsx row count', x.rows.length === 4, String(x.rows.length));
    check('xlsx header', JSON.stringify(x.rows[0]) === JSON.stringify(['ID', 'Name', 'Count', 'Score', 'Active', 'Joined', 'Amount', 'Big', 'Missing']));
    check('xlsx values', x.rows[1][0] === 1 && x.rows[1][1] === 'Alicja żółć' && x.rows[1][3] === 12.5 && x.rows[1][4] === true);
    check('xlsx date', ms(x.rows[1][5]) === Date.UTC(2024, 0, 15, 12, 30, 45), String(ms(x.rows[1][5])));
    check('xlsx NaN as string', x.rows[3][3] === 'NaN', String(x.rows[3][3]));
    check('xlsx spaced string', x.rows[2][6] === '  spaced  ', JSON.stringify(x.rows[2][6]));

    const m = await drainXlsb(path.join(dir, 'rs-multi.xlsb'));
    check('multi sheet names', JSON.stringify(m.names) === JSON.stringify(['data1', 'data2', 'empty']), m.names.join(','));
    check('multi first sheet', JSON.stringify(m.rows) === JSON.stringify([['K', 'V'], ['a', 1], ['b', 2]]), JSON.stringify(m.rows));

    // --- exceljs (independent implementation) over Rust xlsx ---
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(dir, 'rs-basic.xlsx'));
    const ws = wb.getWorksheet('Sheet1');
    check('exceljs sheet found', !!ws);
    if (ws) {
        const header = ws.getRow(1).values as any[];
        check('exceljs header', header[1] === 'ID' && header[2] === 'Name' && header[9] === 'Missing', JSON.stringify(header));
        const r2 = ws.getRow(2).values as any[];
        const d = r2[6] instanceof Date ? (r2[6] as Date).getTime() : null;
        check(
            'exceljs values',
            r2[1] === 1 && r2[2] === 'Alicja żółć' && r2[3] === 42 && r2[4] === 12.5 && r2[5] === true &&
            d === Date.UTC(2024, 0, 15, 12, 30, 45) && r2[7] === 1234.5 && r2[8] === '9007199254740993',
            JSON.stringify(r2.map((v) => (v instanceof Date ? v.toISOString() : v)))
        );
        const r3 = ws.getRow(3).values as any[];
        check('exceljs escapes/bool', r3[2] === 'Bob & <Co> "q" \'x\'' && r3[5] === false, JSON.stringify(r3));
        check('exceljs row count', ws.rowCount === 4, String(ws.rowCount));
        const shortRange = String((ws.dimensions as unknown as { shortRange: string }).shortRange);
        check('exceljs dimension', shortRange === 'A1:I4', shortRange);
    }

    if (failures > 0) {
        console.log(`${failures} FAILURES`);
        process.exit(1);
    }
    console.log('ALL REVERSE CHECKS PASSED');
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
