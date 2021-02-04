import { Renderer } from 'xlsx-renderer';
import * as ExcelJs from 'exceljs';
import {CellTemplateDebugPool} from "xlsx-renderer/lib/CellTemplateDebugPool";


export const VM1: object = {
  firstLineVariable: 'If you see this text, it means that the variable was replaced correctly :)',
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
  public static renderer: Renderer;

  constructor(private templateName: string, private viewModel: any, private buttonId: string) {
    const btn: HTMLElement | null = document.getElementById(buttonId);

    // There is no need to generate multiple renderers,
    // If the renderer hasn't been created yet, we are creating one and storing it in the static field to avoid multiple creation.
    if (!GenerateXLSXFile.renderer) {
      GenerateXLSXFile.renderer = new Renderer(new CellTemplateDebugPool());
    }

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
    console.log('exportXLSX view model:: this.viewModel', this.viewModel);

    try {
      const xlsxBlob: Blob = await this.onRetrieveTemplate();
      const fileReader: FileReader = new FileReader();
      fileReader.readAsArrayBuffer(xlsxBlob);

      fileReader.addEventListener('loadend', async (e: ProgressEvent<FileReader>) => {
        const templateFileBuffer: string | ArrayBuffer | null = fileReader.result;
        if (templateFileBuffer instanceof ArrayBuffer) {
          const result: ExcelJs.Workbook = await GenerateXLSXFile.renderer.renderFromArrayBuffer(templateFileBuffer, this.viewModel);
          const buffer: ExcelJs.Buffer = await result.xlsx.writeBuffer();
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
