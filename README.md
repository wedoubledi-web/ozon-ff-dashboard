# Ozon — короба для фулфилмента

Веб-страница для **Дениса и ФФ**: какой товар в какой короб, куда едет.

## Ссылка

**https://wedoubledi-web.github.io/ozon-ff-dashboard/**

Открывается в браузере на телефоне или компе — Cursor не нужен.

## Что внутри

- Поиск по артикулу → сразу **номер короба**
- Список всех коробов с составом
- Таблица «по товару»
- Кнопка **Скачать Excel**

## Обновление (из Cursor или терминала)

```bash
python3 Бизнес/Маркетплейсы/Ozon/scripts/publish_ff_dashboard.py
```

Или по шагам:

```bash
python3 Бизнес/Маркетплейсы/Ozon/scripts/ozon_ff_boxes.py
python3 Бизнес/Маркетплейсы/Ozon/scripts/publish_ff_dashboard.py --skip-build
```
