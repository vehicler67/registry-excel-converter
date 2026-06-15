#!/bin/bash
# ╔══════════════════════════════════════════════════════════════╗
# ║  등기부등본 PDF → Excel 변환 실행기 (Mac)                       ║
# ║  더블클릭으로 실행 — PDF 선택 → 자동 변환 → Excel 열기            ║
# ╚══════════════════════════════════════════════════════════════╝
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║    등기부등본 PDF → Excel 변환기 v1.0                   ║"
echo "  ║    HanaXellOcr0.7 역공학 기반 — 제2 독립 도구            ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""

# Python3 확인
PYTHON=$(which python3 2>/dev/null)
if [ -z "$PYTHON" ]; then
    echo "❌ Python3가 설치되어 있지 않습니다."
    echo "   https://www.python.org/downloads/ 에서 설치해주세요."
    read -p "아무 키나 누르면 종료됩니다..."
    exit 1
fi

# 의존성 확인
$PYTHON -c "import fitz, PIL, pytesseract, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  필수 라이브러리가 없습니다. 설치를 진행합니다..."
    pip3 install pymupdf pytesseract pillow openpyxl 2>&1
    if [ $? -ne 0 ]; then
        echo ""
        echo "❌ 라이브러리 설치에 실패했습니다."
        echo "   수동 설치: pip3 install pymupdf pytesseract pillow openpyxl"
        read -p "아무 키나 누르면 종료됩니다..."
        exit 1
    fi
    echo "✅ 라이브러리 설치 완료!"
    echo ""
fi

# 변환 실행
$PYTHON "$SCRIPT_DIR/변환기.py"

echo ""
read -p "아무 키나 누르면 종료됩니다..."
