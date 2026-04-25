/**
 * 문서 내보내기 대화상자
 *
 * HWP/HWPX 를 DOCX, Markdown, HTML 형식으로 내보낼 수 있다.
 */
import { ModalDialog } from './dialog';

export type ExportFormat = 'hwp' | 'hwpx' | 'docx' | 'md' | 'html';

interface ExportFormatOption {
  format: ExportFormat;
  label: string;
  extension: string;
  mimeType: string;
}

export const EXPORT_FORMATS: ExportFormatOption[] = [
  { format: 'hwp', label: '한글 문서 (.hwp)', extension: '.hwp', mimeType: 'application/x-hwp' },
  { format: 'hwpx', label: '한글 문서 XML (.hwpx)', extension: '.hwpx', mimeType: 'application/hwp+zip' },
  { format: 'docx', label: 'Word 문서 (.docx)', extension: '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  { format: 'md', label: 'Markdown (.md)', extension: '.md', mimeType: 'text/markdown' },
  { format: 'html', label: 'HTML (.html)', extension: '.html', mimeType: 'text/html' },
];

class ExportDialog extends ModalDialog {
  private defaultName: string;
  private formatOptions: ExportFormatOption[];
  private formatSelect!: HTMLSelectElement;
  private input!: HTMLInputElement;
  private resolve!: (value: { name: string; format: ExportFormat } | null) => void;

  constructor(defaultName: string, formats?: ExportFormatOption[]) {
    super('문서 내보내기', 400);
    this.defaultName = defaultName;
    this.formatOptions = formats ?? EXPORT_FORMATS;
  }

  protected createBody(): HTMLElement {
    const body = document.createElement('div');
    body.style.padding = '16px 20px';

    // 포맷 선택
    const formatLabel = document.createElement('label');
    formatLabel.textContent = '포맷(F):';
    formatLabel.style.display = 'block';
    formatLabel.style.marginBottom = '6px';
    formatLabel.style.fontSize = '13px';
    body.appendChild(formatLabel);

    this.formatSelect = document.createElement('select');
    this.formatSelect.style.width = '100%';
    this.formatSelect.style.boxSizing = 'border-box';
    this.formatSelect.style.height = '26px';
    this.formatSelect.style.padding = '2px 6px';
    this.formatSelect.style.border = '1px solid #b4b4b4';
    this.formatSelect.style.fontSize = '13px';

    this.formatOptions.forEach((opt, idx) => {
      const option = document.createElement('option');
      option.value = opt.format;
      option.textContent = opt.label;
      this.formatSelect.appendChild(option);
    });
    body.appendChild(this.formatSelect);

    // 파일 이름 입력
    const nameLabel = document.createElement('label');
    nameLabel.textContent = '파일 이름(N):';
    nameLabel.style.display = 'block';
    nameLabel.style.marginTop = '12px';
    nameLabel.style.marginBottom = '6px';
    nameLabel.style.fontSize = '13px';
    body.appendChild(nameLabel);

    this.input = document.createElement('input');
    this.input.type = 'text';
    this.input.value = this.defaultName;
    this.input.style.width = '100%';
    this.input.style.boxSizing = 'border-box';
    this.input.style.height = '26px';
    this.input.style.padding = '2px 6px';
    this.input.style.border = '1px solid #b4b4b4';
    this.input.style.fontSize = '13px';

    this.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.onConfirm();
        this.hide();
      }
    });

    // 포맷 변경 시 파일 이름 확장자 자동 업데이트
    this.formatSelect.addEventListener('change', () => {
      const format = this.formatSelect.value as ExportFormat;
      const ext = this.formatOptions.find(f => f.format === format)?.extension ?? '.docx';
      const baseName = this.input.value.replace(/\.[^.]+$/, '');
      this.input.value = baseName + ext;
    });

    body.appendChild(this.input);

    return body;
  }

  protected onConfirm(): void {
    const name = this.input.value.trim();
    const format = this.formatSelect.value as ExportFormat;

    if (!name) return;

    // 파일 이름에 확장자가 없으면 추가
    const ext = this.formatOptions.find(f => f.format === format)?.extension ?? '.docx';
    const fileName = name.toLowerCase().endsWith(ext) ? name : name + ext;

    this.resolve({ name: fileName, format });
  }

  override hide(): void {
    this.resolve(null);
    super.hide();
  }

  showAsync(): Promise<{ name: string; format: ExportFormat } | null> {
    return new Promise((resolve) => {
      let resolved = false;
      this.resolve = (v: { name: string; format: ExportFormat } | null) => {
        if (!resolved) {
          resolved = true;
          resolve(v);
        }
      };
      super.show();
      requestAnimationFrame(() => {
        this.input.focus();
        this.input.select();
      });
    });
  }
}

/** 내보내기 대화상자를 표시하고 사용자가 선택한 파일 이름과 포맷을 반환한다. 취소 시 null. */
export function showExport(defaultName: string, formats?: ExportFormatOption[]): Promise<{ name: string; format: ExportFormat } | null> {
  return new ExportDialog(defaultName, formats).showAsync();
}

/** 내보내기 대화상자를 표시한다. HWP/HWPX 포맷은 제외되며, DOCX/MD/HTML 만 선택 가능하다. */
export function showExportDialog(defaultName: string): Promise<{ name: string; format: ExportFormat } | null> {
  const exportFormats = EXPORT_FORMATS.filter(f => !['hwp', 'hwpx'].includes(f.format));
  return new ExportDialog(defaultName, exportFormats).showAsync();
}