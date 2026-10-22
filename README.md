# Підбір подарунка · Geekshop Ukraine

Сторінка-підбирач подарунків для клієнтів geekshopukraine.com.ua.

- `index.html` — сама сторінка (GitHub Pages).
- `catalog-fallback.json` — запасний знімок каталогу, якщо живий каталог недоступний.
- `apps-script/` — копія коду Google Apps Script (таблиця «Заявки підбирача Geekshop»):
  записує заявки і раз на 12 годин оновлює каталог з Prom API. Токен Prom зберігається
  у властивостях скрипту, не в коді.
