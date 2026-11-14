# TabEscape

Проект находится в [tabescape_app](tabescape_app/README.md).
Материал для пяти минут питча и сценарий показа — [PITCH.md](PITCH.md).
Настройки развёртывания Flask на Render — [render.yaml](render.yaml).

Интерфейс из `TabEscape (2).html` подключён к Flask на `/`. Исходный HTML сохранён как самостоятельный референс; рабочие файлы сайта находятся в `tabescape_app/templates/` и `tabescape_app/static/`. Ходы используют серверные `/api/start` и `/api/action`.

В библиотеке теперь 8 бесплатных историй. У всех четырёх масок есть окно подготовки партии; Premium и Endless доступны как демо. Подключение OpenAI описано в [инструкции](tabescape_app/README.md#premium-и-openai), шаблон переменных — [.env.example](.env.example). Копировать его нужно в `tabescape_app/.env`; настоящий ключ в Git не добавлять.
