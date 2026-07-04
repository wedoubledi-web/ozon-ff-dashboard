# Ozon — короба для фулфилмента

Веб-страница для **Дениса и ФФ**: какой товар в какой короб, куда едет.

## Ссылка

**https://wedoubledi-web.github.io/ozon-ff-dashboard/**

Открывается в браузере на телефоне или компе — Cursor не нужен.

## Что внутри

1. **Вписать номер поставки** (например `116057614-1`) → «Открыть»
2. Или нажать на чип с недавней поставкой
3. Поиск по артикулу → номер короба
4. **Скачать Excel** именно по этой поставке

## Обновление

**Автоматически:** GitHub Actions каждые 2 часа тянет данные из Ozon API и публикует сайт — Mac может быть выключен.

**Вручную** (из Cursor или терминала):

```bash
python3 Бизнес/Маркетплейсы/Ozon/scripts/publish_ff_dashboard.py
```

Или по шагам:

```bash
python3 Бизнес/Маркетплейсы/Ozon/scripts/ozon_ff_boxes.py
python3 Бизнес/Маркетплейсы/Ozon/scripts/publish_ff_dashboard.py --skip-build
```

Принудительный запуск в облаке: GitHub → `kursor-wb` → Actions → **Ozon FF Dashboard** → Run workflow.
