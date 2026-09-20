/**
 * Generates parity fixtures with the TypeScript reference implementation.
 *
 * Run from anywhere with the TS repo's ts-node:
 *   node_modules/.bin/ts-node --transpile-only --skip-project \
 *     --compiler-options '{"module":"commonjs"}' gen_fixtures.ts <outDir>
 *
 * All data is deterministic (no Math.random / wall-clock dates).
 */
import * as path from 'path';
import * as fs from 'fs';
import { XlsbWriter } from '../../justybase_spreadsheet_tasks/src/XlsbWriter';
import { XlsxWriter } from '../../justybase_spreadsheet_tasks/src/XlsxWriter';
import { F } from '../../justybase_spreadsheet_tasks/src/Formats';

const outDir = process.argv[2] ?? path.join(__dirname, '..', 'tests', 'fixtures');
fs.mkdirSync(outDir, { recursive: true });

const HEADERS = ['ID', 'Name', 'Count', 'Score', 'Active', 'Joined', 'Amount', 'Big', 'Missing'];

const ROWS: any[][] = [
    [
        1,
        'Alicja żółć',
        42,
        12.5,
        true,
        new Date(Date.UTC(2024, 0, 15, 12, 30, 45)),
        { value: 1234.5, format: F.TWO_DECIMALS },
        BigInt('9007199254740993'),
        null,
    ],
    [
        -7,
        'Bob & <Co> "q" \'x\'',
        1 << 29, // beyond RK range -> double path
        -0.001,
        false,
        { value: new Date(Date.UTC(2023, 11, 31, 8, 0, 0)), format: F.DATE_ISO },
        '  spaced  ',
        BigInt(0),
        undefined,
    ],
    [
        0,
        '',
        -(1 << 29), // RK lower boundary (inclusive)
        Number.NaN, // writers store non-finite numbers as strings
        true,
        new Date(Date.UTC(1899, 11, 30, 0, 0, 0)), // OA epoch itself
        { value: 5, format: F.CURRENCY_PLN },
        BigInt('-123456789012345678901234567890'),
        null,
    ],
];

async function main(): Promise<void> {
    for (const ext of ['xlsb', 'xlsx']) {
        const basicPath = path.join(outDir, `ts-basic.${ext}`);
        if (ext === 'xlsb') {
            const w = new XlsbWriter(basicPath);
            w.addSheet('Sheet1');
            w.writeSheet(ROWS, HEADERS);
            await w.finalize();
        } else {
            const w = new XlsxWriter(basicPath);
            w.addSheet('Sheet1');
            w.writeSheet(ROWS, HEADERS);
            await w.finalize();
        }

        const multiPath = path.join(outDir, `ts-multi.${ext}`);
        if (ext === 'xlsb') {
            const w = new XlsbWriter(multiPath);
            w.addSheet('data1');
            w.writeSheet(
                [
                    ['a', 1],
                    ['b', 2],
                ],
                ['K', 'V']
            );
            w.addSheet('data2', true);
            w.writeSheet([['only', false]], ['X', 'Y']);
            w.addSheet('empty');
            w.writeSheet([], ['H1', 'H2'], false);
            await w.finalize();
        } else {
            const w = new XlsxWriter(multiPath);
            w.addSheet('data1');
            w.writeSheet(
                [
                    ['a', 1],
                    ['b', 2],
                ],
                ['K', 'V']
            );
            w.addSheet('data2', true);
            w.writeSheet([['only', false]], ['X', 'Y']);
            w.addSheet('empty');
            w.writeSheet([], ['H1', 'H2'], false);
            await w.finalize();
        }
        console.log(`wrote ${basicPath} + ${multiPath}`);
    }
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
