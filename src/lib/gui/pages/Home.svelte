<script lang="ts">
    import Excel from "exceljs";
    import {
        TextReader,
        BlobWriter,
        ZipWriter,
        BlobReader,
    } from "@zip.js/zip.js";
    import Swal from "sweetalert2";
    import z from "zod";

    type JsonWorksheet = {
        name: string;
        rows: Record<string, string>[];
    };

    type BlobInfo = {
        blob: Blob;
        name: string;
    };

    const title: string = "Excel-2-JSON";

    let inputFileExcel = $state<HTMLInputElement>();
    let inputFileJson = $state<HTMLInputElement>();

    export async function sleep(ms: number): Promise<void> {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    async function handleExcelFiles(): Promise<void> {
        try {
            Swal.fire({
                didOpen: () => {
                    Swal.showLoading();
                },
            });

            const files = [...(inputFileExcel?.files ?? [])];

            if (files.length <= 0) {
                console.error("no files to handle");
                return;
            }

            const jsonWorksheets: JsonWorksheet[] = [];

            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const stream = await file.arrayBuffer();
                const workbook = new Excel.Workbook();
                await workbook.xlsx.load(stream);
                const newJsonWorksheets = await workbookToJson(workbook);

                for (const jsonWorksheet of newJsonWorksheets) {
                    const baseName = file.name.split(".")[0] ?? "file";
                    const worksheetName = jsonWorksheet.name;
                    jsonWorksheet.name = baseName;

                    if (newJsonWorksheets.length > 1) {
                        jsonWorksheet.name = `${baseName} - ${worksheetName}`;
                    }

                    jsonWorksheets.push(jsonWorksheet);
                }
            }

            const blobInfo = await jsonWorksheetsToBlob(jsonWorksheets);
            const zipBlobUrl = window.URL.createObjectURL(blobInfo.blob);
            const anchor = document.createElement("a");
            anchor.target = "_blank";
            anchor.href = zipBlobUrl;
            anchor.download = blobInfo.name;
            anchor.click();
        } catch (error) {
            console.error("failed to handle excel files", error);
        } finally {
            Swal.close();
        }
    }

    async function jsonWorksheetsToBlob(
        jsonWorksheets: JsonWorksheet[]
    ): Promise<BlobInfo> {
        const blobInfo: BlobInfo = {
            blob: new Blob(),
            name: "",
        };

        if (jsonWorksheets.length === 1) {
            const jsonWorksheet = jsonWorksheets[0];
            const jsonString = JSON.stringify(jsonWorksheet.rows, null, 4);
            blobInfo.blob = new Blob([jsonString], {
                type: "application/json",
            });

            blobInfo.name = jsonWorksheet.name + ".json";
            return blobInfo;
        }

        const zipFileWriter = new BlobWriter();
        const zipWriter = new ZipWriter(zipFileWriter, {
            compressionMethod: 0,
            level: 0,
        });

        for (const jsonWorksheet of jsonWorksheets) {
            const rows = jsonWorksheet.rows;
            const jsonString = JSON.stringify(rows, null, 4);
            const blobReader = new TextReader(jsonString);
            const jsonFileName = jsonWorksheet.name + ".json";
            zipWriter.add(jsonFileName, blobReader);
        }

        blobInfo.name = "json.zip";
        blobInfo.blob = await zipWriter.close();
        return blobInfo;
    }

    async function handleJsonFiles(): Promise<void> {
        try {
            Swal.fire({
                didOpen: () => {
                    Swal.showLoading();
                },
            });

            const files = [...(inputFileJson?.files ?? [])];

            if (files.length <= 0) {
                console.error("no files to handle");
                return;
            }

            const blobsInfo: BlobInfo[] = [];

            for (const file of files) {
                const rawJson = await file.text();
                const rows: unknown[] = z
                    .unknown()
                    .array()
                    .catch([])
                    .parse(JSON.parse(rawJson));

                const records: string[][] = [];
                const cols: string[] = [];

                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    const record: string[] = [];

                    switch (typeof row) {
                        case "object": {
                            if (row == null) {
                                record.push(`${row}`);
                                break;
                            }

                            const keys = Object.keys(row);

                            if (i === 0) {
                                cols.push(...keys);
                                records.push(keys);
                            }

                            const parsedRow = z
                                .record(z.string(), z.unknown())
                                .parse(row);

                            for (const col of cols) {
                                const value = `${parsedRow[col] ?? ""}`;
                                record.push(value);
                            }
                            break;
                        }

                        default: {
                            record.push(`${row}`);
                            break;
                        }
                    }

                    records.push(record);
                }

                const workbook = jsonToExcel(records);
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer]);
                const baseName = file.name.split(".")[0] ?? "file";

                const blobInfo: BlobInfo = {
                    blob: blob,
                    name: baseName,
                };

                blobsInfo.push(blobInfo);
            }

            const blobInfo = await excelBlobsToBlob(blobsInfo);
            const blobUrl = window.URL.createObjectURL(blobInfo.blob);
            const anchor = document.createElement("a");
            anchor.target = "_blank";
            anchor.href = blobUrl;
            anchor.download = blobInfo.name;
            anchor.click();
        } catch (error) {
            console.error("failed to handle json files", error);
        } finally {
            Swal.close();
        }
    }

    async function excelBlobsToBlob(blobsInfo: BlobInfo[]): Promise<BlobInfo> {
        const blobInfo: BlobInfo = {
            blob: new Blob(),
            name: "",
        };

        if (blobsInfo.length === 1) {
            const blobInfo = blobsInfo[0];
            blobInfo.name = blobInfo.name + ".xlsx";
            return blobInfo;
        }

        const zipFileWriter = new BlobWriter();
        const zipWriter = new ZipWriter(zipFileWriter, {
            compressionMethod: 0,
            level: 0,
        });

        for (const blobInfo of blobsInfo) {
            const blobReader = new BlobReader(blobInfo.blob);
            const jsonFileName = blobInfo.name + ".xlsx";
            zipWriter.add(jsonFileName, blobReader);
        }

        blobInfo.name = "excel.zip";
        blobInfo.blob = await zipWriter.close();
        return blobInfo;
    }

    function jsonToExcel(rows: string[][]): Excel.Workbook {
        const workbook = new Excel.Workbook();
        const worksheet = workbook.addWorksheet("Sheet 1");
        worksheet.addRows(rows);
        return workbook;
    }

    async function workbookToJson(
        workbook: Excel.Workbook
    ): Promise<JsonWorksheet[]> {
        const jsonWorksheets: JsonWorksheet[] = [];

        const sheetNames: string[] = [];
        workbook.eachSheet((worksheet) => {
            sheetNames.push(worksheet.name);
        });

        for (let i = 0; i < sheetNames.length; i++) {
            const sheetName = sheetNames[i];
            const sheet = workbook.getWorksheet(sheetName);
            if (sheet == null) {
                continue;
            }

            const sheetJson = await sheetToJson(sheet);
            jsonWorksheets.push(sheetJson);
        }

        return jsonWorksheets;
    }

    async function sheetToJson(
        worksheet: Excel.Worksheet
    ): Promise<JsonWorksheet> {
        const jsonWorksheet: JsonWorksheet = {
            name: worksheet.name,
            rows: [],
        };

        const cols: string[] = [];
        {
            const row: Excel.Row = worksheet.getRow(1);

            for (let i = 1; i <= row.cellCount; i++) {
                const cell: string = row.getCell(i).toString();
                if (cell == "") {
                    break;
                }

                cols.push(cell);
            }
        }

        for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
            const row: Excel.Row = worksheet.getRow(rowIndex);
            const rowData: Record<string, string> = {};

            for (let colIndex = 1; colIndex <= row.cellCount; colIndex++) {
                const cell: string = row.getCell(colIndex).toString();
                const key = cols[colIndex - 1] ?? "";
                if (key == "") {
                    continue;
                }

                rowData[key] = cell;
            }

            jsonWorksheet.rows.push(rowData);
        }

        return jsonWorksheet;
    }
</script>

<svelte:head>
    <title>{title}</title>
</svelte:head>

<section>
    <h1 class="mb-4">{title}</h1>

    <div class="file is-fullwidth mb-2">
        <label class="file-label">
            <input
                bind:this={inputFileExcel}
                multiple
                class="file-input"
                type="file"
                name="resume"
                accept="application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onchange={() => handleExcelFiles()}
            />
            <span class="file-cta">
                <span class="file-icon">
                    <i class="fas fa-upload"></i>
                </span>
                <span class="file-label"> Excel to JSON </span>
            </span>
        </label>
    </div>

    <div class="file is-fullwidth mb-2">
        <label class="file-label">
            <input
                bind:this={inputFileJson}
                multiple
                class="file-input"
                type="file"
                name="resume"
                accept="application/json"
                onchange={() => handleJsonFiles()}
            />
            <span class="file-cta">
                <span class="file-icon">
                    <i class="fas fa-upload"></i>
                </span>
                <span class="file-label"> JSON to Excel </span>
            </span>
        </label>
    </div>
</section>

<style>
    section {
        display: flex;
        width: 100%;
        flex-direction: column;
    }

    h1 {
        width: 100%;
    }

    .file {
        margin: 0;
    }

    .file-cta {
        width: 100%;
    }
</style>
