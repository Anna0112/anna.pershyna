# CV Generator (реальная версия под ваш макет)

Минимальная рабочая версия для ПК: берет PDF-резюме, извлекает текст и собирает **styled DOCX** под ваш референс (шапка + подложка + секции).

## 1) Установка

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## 2) Подготовка файлов

- Положите PDF резюме: `data/anna_cv.pdf`
- Положите шапку (из вашего изображения): `templates/header.png`
- Положите подложку/лого (оранжевый знак): `templates/watermark.png`

> Шрифт `Atkinson Hyperlegible` должен быть установлен в системе, иначе Word подставит замену.

## 3) Генерация

```bash
python app/run.py \
  --pdf data/anna_cv.pdf \
  --output output/Anna_Pershyna_CV.docx \
  --header-image templates/header.png \
  --watermark-image templates/watermark.png \
  --context-json output/context.json
```

## Что уже адаптировано под вашу структуру

- Шапка документа через отдельный PNG.
- Подложка/бренд-элемент в нижней части страницы.
- Секции: `Summary`, `Skills`, `Professional Experience`, `Education & Certifications`.
- Базовая авто-нарезка текста из PDF по секциям.

## Ограничение MVP

PDF-парсинг всегда “грязный”, поэтому после первой генерации откройте `output/context.json` и при необходимости поправьте блоки (особенно experience), затем можно доработать парсер под ваш стабильный формат CV.
