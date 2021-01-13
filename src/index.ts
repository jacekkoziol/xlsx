import { Renderer } from '../node_modules/xlsx-renderer';
import * as Excel from '../node_modules/exceljs';

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

export class GenerateXLSXFile {
  constructor(private templateName: string, private viewModel: any, private buttonId: string) {
    const btn: HTMLElement | null = document.getElementById(buttonId);
    console.log('Init');

    if (btn) {
      btn.addEventListener('click', () => {
        console.log(`Button ID: ${this.buttonId} clicked`);
        this.exportXLSX()
      }, false)
    }
  }

  public async onRetrieveTemplate(): Promise<Blob> {
    // return fetch('./xlsx-templates/template.xlsx').then((r: Response) => r.blob());
    return fetch(`./xlsx-templates/${this.templateName}`).then((r: Response) => r.blob());
  }

  public exportXLSX(): void {
    console.log('exportXLSX view model:: this.viewModel');
    this.onRetrieveTemplate()
      .then((xlsxBlob: Blob) => {
        const reader: FileReader = new FileReader();
        reader.readAsArrayBuffer(xlsxBlob);
        reader.addEventListener('loadend', async (e: ProgressEvent<FileReader>) => {
          const renderer: Renderer = new Renderer();

          if (reader.result instanceof ArrayBuffer) {
            const workbook: Excel.Workbook = new Excel.Workbook();
            await workbook.xlsx.load(reader.result);
            const result: Excel.Workbook = await renderer.render(() => Promise.resolve(workbook), this.viewModel);

            await result.xlsx
              .writeBuffer()
              .then((buffer: Excel.Buffer) => {
                this.saveBlobToFile(new Blob([buffer]), `${Date.now()}_result_report.xlsx`);
              })
              .catch((err: Error) => console.log('Error writing excel export', err));
          }
        });
      })
      .catch((err: Error) => console.log('Error:', err));
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
