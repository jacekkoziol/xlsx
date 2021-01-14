import { Renderer } from 'xlsx-renderer';
import * as Excel from 'exceljs';
import {CellTemplateDebugPool} from "xlsx-renderer/lib/CellTemplateDebugPool";

export const VM1: object = {
  projects: [
    {
      name: 'ExcelJS',
      role: 'maintainer',
      platform: 'github',
      link: 'https://github.com/exceljs/exceljs',
      stars: 5300,
      forks: 682,
    },
    {
      name: 'xlsx-import',
      role: 'owner',
      platform: 'github',
      link: 'https://github.com/siemienik/xlsx-import',
      stars: 2,
      forks: 0,
    },
    {
      name: 'xlsx-import',
      role: 'owner',
      platform: 'npm',
      link: 'https://www.npmjs.com/package/xlsx-import',
      stars: 'n.o.',
      forks: 'n.o.',
    },
    {
      name: 'xlsx-renderer',
      role: 'owner',
      platform: 'github',
      link: 'https://github.com/siemienik/xlsx-renderer',
      stars: 1,
      forks: 0,
    },
    {
      name: 'xlsx-renderer',
      role: 'owner',
      platform: 'npm',
      link: 'https://www.npmjs.com/package/xlsx-renderer',
      stars: 'n.o.',
      forks: 'n.o.',
    },
    {
      name: 'TS Package Structure',
      role: 'owner',
      platform: 'github',
      link: 'https://github.com/Siemienik/ts-package-structure',
      stars: 2,
      forks: 0,
    },
  ],
};

const VM_GT: object = {
  main: 'Main Text',
  totalSum: 291827627,
  rows: [
    { id: 1, tier: 1, name: 'Agency name', totalAmount: 1000 },
    { id: 2, tier: 1, name: 'Agency name 2', totalAmount: 1980 },
    { id: 3, tier: 4, name: 'Agency name 3', totalAmount: 1283710 },
    { id: 4, tier: 3, name: 'Agency name 4', totalAmount: 2000000 },
  ],
};

export class GenerateXLSXFile {
  private renderer: Renderer;

  constructor(private templateName: string, private viewModel: any, private buttonId: string) {
    const btn: HTMLElement | null = document.getElementById(buttonId);
    // todo @siemienik, add information about Debug possibilities into Readme
    // todo @siemienik, add information about extending CellTemplatePool.
    // todo @siemienik, add logging cell address into console.log inside CellTemplateDebugPool.match();
    this.renderer = new Renderer(new CellTemplateDebugPool());

    console.log('Init');

    if (btn) {
      btn.addEventListener('click', () => {
        console.log(`Button ID: ${this.buttonId} clicked`);
        this.exportXLSX()
      }, false)
    }
  }

  public async onRetrieveTemplate(): Promise<Blob> {
    return fetch(`./xlsx-templates/${this.templateName}`).then((r: Response) => r.blob());
  }

  public async exportXLSX(): Promise<void> {
    console.log('exportXLSX view model:: this.viewModel');
    try {
      const xlsxBlob: Blob = await this.onRetrieveTemplate();
      const fileReader: FileReader = new FileReader();
      fileReader.readAsArrayBuffer(xlsxBlob);

      fileReader.addEventListener('loadend', async (e: ProgressEvent<FileReader>) => {
        const templateFileBuffer = fileReader.result;
        if (templateFileBuffer instanceof ArrayBuffer) {
          // todo @siemeinik, Add information about correct template factory
          // todo @siemienik, Add feature which detect that template and output objects is same (Consider about thrown an error or warning, or do clone if possible and works properly)
          // todo @siemienik, Add possibility to load from fileBuffer.
          const templateFactory: () => Promise<Excel.Workbook> = () => { // All this logic must be provided into xlsx-renderer as a function
            const workbook: Excel.Workbook = new Excel.Workbook();
            return workbook.xlsx.load(templateFileBuffer);
          };

          const result: Excel.Workbook = await this.renderer.render(templateFactory, this.viewModel);
          const buffer: Excel.Buffer = await result.xlsx.writeBuffer()
          this.saveBlobToFile(new Blob([buffer]), `${Date.now()}_result_report.xlsx`);
        }
      });
    } catch (err) {
      console.log('Error:', err);
    }
  }

  // Utilities - File Save
  // ---------------------------------------------------------------------------
  private saveBlobToFile(blob: Blob, fileName: string = 'File.xlsx'): void {
    const link: HTMLAnchorElement = document.createElement('a');
    const url: string = window.URL.createObjectURL(blob);
    link.href = url;
    link.download = fileName;
    link.target = '_blank';
    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => {
      window.URL.revokeObjectURL(url);
    }, 4000);
  }
}

// Initialize
// -----------------------------------------------------------------------------
new GenerateXLSXFile('template.xlsx', VM1, 'exportFile1');
new GenerateXLSXFile('template-hyperlink.xlsx', VM1, 'exportFileHyperlink');
