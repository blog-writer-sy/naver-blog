"""원고(.txt)의 (★표-N) 표를 엑셀로 뽑는다.

사용법: python3 도구/표_엑셀.py 원고1.txt [원고2.txt ...] [-o 출력.xlsx]
- 원고가 하나면 출력은 원고와 같은 폴더의 '{원고이름}_표.xlsx'
- 표마다 시트 1개 (원고 여러 개면 '원고1_표-1'처럼 앞에 번호)
- 서식: 첫 줄 진한 색 배경 + 흰 굵은 글씨, 아래 줄 글씨는 표 색, 테두리 연회색 굵게
- 한 원고 안의 표는 색을 다르게 (남색, 진녹, 갈색, 보라, 청록 순)
"""
import re, sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

PAL = ['1F4E79', '375623', '833C0B', '5B3A7E', '1E6B6B']
FONT = '맑은 고딕'
SIDE = Side(style='medium', color='BFBFBF')
BORDER = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)

def tables(text):
    for m in re.finditer(r'\(★(표-\d+)\)\n((?:\|.*\|\n?)+)', text):
        rows = [[c.strip() for c in l.strip().strip('|').split('|')] for l in m.group(2).strip().split('\n')]
        yield m.group(1), [r for r in rows if not re.fullmatch(r'[-:| ]+', '|'.join(r))]

def main(argv):
    out = None
    if '-o' in argv:
        i = argv.index('-o'); out = argv[i + 1]; argv = argv[:i] + argv[i + 2:]
    if not argv:
        sys.exit(__doc__)
    if out is None:
        out = os.path.splitext(argv[0])[0] + '_표.xlsx'
    wb = Workbook(); wb.remove(wb.active)
    for n, path in enumerate(argv, 1):
        text = open(path, encoding='utf-8').read()
        for k, (name, rows) in enumerate(tables(text)):
            color = PAL[k % len(PAL)]
            ws = wb.create_sheet(name if len(argv) == 1 else f'원고{n}_{name}')
            ws.sheet_properties.tabColor = color
            ws.sheet_view.showGridLines = False
            widths = {}
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    c = ws.cell(i + 1, j + 1, v)
                    c.border = BORDER
                    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    if i == 0:
                        c.fill = PatternFill('solid', fgColor=color)
                        c.font = Font(name=FONT, bold=True, size=10, color='FFFFFF')
                    else:
                        c.font = Font(name=FONT, size=10, color=color)
                    widths[j] = max(widths.get(j, 0), sum(2 if ord(ch) > 127 else 1 for ch in v))
                ws.row_dimensions[i + 1].height = 24
            for j, w in widths.items():
                ws.column_dimensions[chr(65 + j)].width = min(max(w + 4, 14), 46)
    if not wb.sheetnames:
        sys.exit('표가 없습니다: (★표-N) 다음 줄에 | 표 | 형식으로 적어 주세요.')
    wb.save(out)
    print(out, wb.sheetnames)

if __name__ == '__main__':
    main(sys.argv[1:])
