/**
 * Benchmarks for the TypeScript reference implementation.
 *
 * Deterministic dataset (no Math.random / wall-clock), same shape as the
 * Rust `parity_bench` example so numbers are comparable. Batch mode
 * (`addSheet` + `writeSheet` + `finalize`), like the README figures.
 *
 * Usage: ts-node ... bench_ts.ts <outDir> [rows] [iters]
 * Prints one JSON document to stdout.
 */
import * as path from 'path';
import * as fs from 'fs';
import { XlsbWriter } from '../../justybase_spreadsheet_tasks/src/XlsbWriter';
import { XlsxWriter } from '../../justybase_spreadsheet_tasks/src/XlsxWriter';
import { XlsbReader } from '../../justybase_spreadsheet_tasks/src/XlsbReader';
import { XlsxReader } from '../../justybase_spreadsheet_tasks/src/XlsxReader';

const outDir = process.argv[2] ?? '/tmp/bench';
const ROWS = parseInt(process.argv[3] ?? '5000', 10);
const ITERS = parseInt(process.argv[4] ?? '3', 10);
const BASE = Date.UTC(2024, 0, 15, 12, 0, 0);

const HEADERS = ['ID', 'Name', 'Count', 'Score', 'Date', 'Active', 'Description'];
const DESC_TAIL = ' z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.';

function buildData(): any[][] {
    const data: any[][] = new Array(ROWS);
    for (let i = 0; i < ROWS; i++) {
        data[i] = [
            i,
            `Produkt ${i} żółć`,
            (i * 7919) % 10000,
            ((i * 104729) % 100000) / 1000,
            new Date(BASE + i * 1000),
            i % 2 === 0,
            `Opis produktu ${i}${DESC_TAIL}`,
        ];
    }
    return data;
}

function median(xs: number[]): number {
    const s = [...xs].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
}

async function benchWriteXlsb(file: string, data: any[][]): Promise<number[]> {
    const times: number[] = [];
    for (let k = 0; k < ITERS; k++) {
        if (fs.existsSync(file)) fs.unlinkSync(file);
        const t0 = performance.now();
        const w = new XlsbWriter(file);
        w.addSheet('Benchmark');
        w.writeSheet(data, HEADERS);
        await w.finalize();
        times.push(performance.now() - t0);
    }
    return times;
}

async function benchWriteXlsx(file: string, data: any[][]): Promise<number[]> {
    const times: number[] = [];
    for (let k = 0; k < ITERS; k++) {
        if (fs.existsSync(file)) fs.unlinkSync(file);
        const t0 = performance.now();
        const w = new XlsxWriter(file);
        w.addSheet('Benchmark');
        w.writeSheet(data, HEADERS);
        await w.finalize();
        times.push(performance.now() - t0);
    }
    return times;
}

async function benchReadXlsb(file: string): Promise<number[]> {
    const times: number[] = [];
    for (let k = 0; k < ITERS; k++) {
        const t0 = performance.now();
        const r = new XlsbReader();
        await r.open(file);
        let acc = 0;
        let n = 0;
        while (await r.read()) {
            for (let i = 0; i < r.fieldCount; i++) {
                const v = r.getValue(i);
                if (typeof v === 'number') acc += v;
                else if (typeof v === 'string') acc += v.length;
                else if (typeof v === 'boolean') acc += v ? 1 : 0;
                else if (v instanceof Date) acc += v.getTime() % 1000;
            }
            n++;
        }
        times.push(performance.now() - t0);
        if (acc === -1 || n === -1) console.log('unreachable');
    }
    return times;
}

async function benchReadXlsx(file: string): Promise<number[]> {
    const times: number[] = [];
    for (let k = 0; k < ITERS; k++) {
        const t0 = performance.now();
        const r = new XlsxReader();
        await r.open(file);
        let acc = 0;
        let n = 0;
        while (await r.read()) {
            for (let i = 0; i < r.fieldCount; i++) {
                const v = r.getValue(i);
                if (typeof v === 'number') acc += v;
                else if (typeof v === 'string') acc += v.length;
                else if (typeof v === 'boolean') acc += v ? 1 : 0;
                else if (v instanceof Date) acc += v.getTime() % 1000;
            }
            n++;
        }
        await r.close();
        times.push(performance.now() - t0);
        if (acc === -1 || n === -1) console.log('unreachable');
    }
    return times;
}

async function main(): Promise<void> {
    fs.mkdirSync(outDir, { recursive: true });
    const data = buildData();
    const xlsb = path.join(outDir, 'bench-ts.xlsb');
    const xlsx = path.join(outDir, 'bench-ts.xlsx');

    const wXlsb = await benchWriteXlsb(xlsb, data);
    const wXlsx = await benchWriteXlsx(xlsx, data);
    const rXlsb = await benchReadXlsb(xlsb);
    const rXlsx = await benchReadXlsx(xlsx);

    const result = {
        impl: 'typescript',
        rows: ROWS,
        iters: ITERS,
        write_xlsb_ms: wXlsb,
        write_xlsx_ms: wXlsx,
        read_xlsb_ms: rXlsb,
        read_xlsx_ms: rXlsx,
        write_xlsb_median_ms: median(wXlsb),
        write_xlsx_median_ms: median(wXlsx),
        read_xlsb_median_ms: median(rXlsb),
        read_xlsx_median_ms: median(rXlsx),
        size_xlsb: fs.statSync(xlsb).size,
        size_xlsx: fs.statSync(xlsx).size,
    };
    console.log(JSON.stringify(result));
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
