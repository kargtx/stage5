/**
 * Скрипт Google Apps Script для автоматического создания двух форм и таблицы
 * Дисциплина: СТАУ, Лабораторная работа №4, Вариант 17 (Тема 2)
 * Студенты: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С
 *
 * ИНСТРУКЦИЯ ПО ЗАПУСКУ (занимает 1 минуту):
 * 1. Откройте в браузере сайт: https://script.google.com
 * 2. Нажмите синюю кнопку «Новый проект» (+ New project)
 * 3. Удалите весь существующий код в окне и вставьте содержимое этого файла
 * 4. Нажмите кнопку «Сохранить» (иконка дискеты), затем «Выполнить» (Run)
 * 5. При первом запуске Google запросит стандартное разрешение на доступ к вашему Google Диску:
 *    нажмите «Просмотреть разрешения» -> выберите свой аккаунт -> «Дополнительно» -> «Перейти на страницу (небезопасно)» -> «Разрешить»
 * 6. В окне «Журнал выполнения» (Execution log) внизу появятся готовые рабочие ссылки на обе формы!
 */

function createAllFormsAndSpreadsheet() {
  Logger.log("=== НАЧАЛО СОЗДАНИЯ ФОРМ ===");

  // ==========================================
  // 1. СОЗДАНИЕ ОПРОСНИКА (АНКЕТА СОЦСЕТИ)
  // ==========================================
  var poll = FormApp.create("Анкета пользователя социальной сети (ЭАС-514С, Вар. 17)");
  poll.setDescription("Опрос проводится в рамках лабораторной работы №4 по дисциплине СТАУ.\nАвторы: Гайнанов Д.И., Гарифуллин К.Р. (гр. ЭАС-514С).");

  // Регистрационные поля
  poll.addTextItem().setTitle("ФИО опрашиваемого").setRequired(true);
  poll.addTextItem().setTitle("E-mail опрашиваемого").setRequired(true);
  poll.addDateItem().setTitle("Дата заполнения анкеты").setRequired(true);

  // 10 вопросов (ровно 2 вопроса Да/Нет)
  poll.addMultipleChoiceItem()
    .setTitle("1. Какую социальную сеть вы посещаете чаще всего?")
    .setChoiceValues(["Telegram", "VK (ВКонтакте)", "Rutube / YouTube", "Другая"])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("2. Сколько времени в среднем в день вы проводите в социальных сетях?")
    .setChoiceValues(["Менее 1 часа", "1–3 часа", "3–5 часов", "Более 5 часов"])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("3. Какова основная цель использования социальной сети?")
    .setChoiceValues([
      "Общение с друзьями и близкими",
      "Просмотр новостей и информационных лент",
      "Учеба и профессиональная деятельность",
      "Развлекательный контент (видео, мемы)",
      "Покупки и поиск услуг"
    ])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("4. С какого устройства вы чаще всего заходите в соцсеть?")
    .setChoiceValues(["Смартфон", "Ноутбук / ПК", "Планшет"])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("5. Какой формат контента вы предпочитаете потреблять?")
    .setChoiceValues([
      "Короткие вертикальные видео (Клипы / Shorts)",
      "Текстовые посты и статьи",
      "Фотографии и подборки",
      "Длинные видео и подкасты"
    ])
    .setRequired(true);

  // Вопрос Да/Нет № 1
  poll.addMultipleChoiceItem()
    .setTitle("6. Публикуете ли вы собственный авторский контент на личной странице?")
    .setChoiceValues(["Да", "Нет"])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("7. Как вы относитесь к персонализированной контекстной рекламе в ленте?")
    .setChoiceValues([
      "Положительно (нахожу полезные товары)",
      "Нейтрально (не обращаю внимания)",
      "Отрицательно (раздражает, блокирую)"
    ])
    .setRequired(true);

  // Вопрос Да/Нет № 2
  poll.addMultipleChoiceItem()
    .setTitle("8. Используете ли вы двухфакторную аутентификацию (2FA) для защиты профиля?")
    .setChoiceValues(["Да", "Нет"])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("9. Состоите ли вы в тематических студенческих или профессиональных сообществах?")
    .setChoiceValues([
      "Да, активно участвую в обсуждениях",
      "Состою только ради объявлений и новостей",
      "Не состою"
    ])
    .setRequired(true);

  poll.addMultipleChoiceItem()
    .setTitle("10. Оцените степень доверия к безопасности хранения персональных данных в соцсетях:")
    .setChoiceValues([
      "Высокая (полностью доверяю)",
      "Средняя (выкладываю минимум данных)",
      "Низкая (не доверяю, опасаюсь утечек)"
    ])
    .setRequired(true);

  Logger.log("--> Опросник создан!");
  Logger.log("Ссылка для респондентов: " + poll.getPublishedUrl());
  Logger.log("Ссылка для редактирования: " + poll.getEditUrl());

  // ==========================================
  // 2. СОЗДАНИЕ ТЕСТА (ТЕСТ ПО ФИЗИКЕ)
  // ==========================================
  var test = FormApp.create("Тест на знание физики (ЭАС-514С, Вар. 17)");
  test.setDescription("Тест проводится в рамках лабораторной работы №4 по дисциплине СТАУ.\nКритерии: 0-5 неуд, 6-7 удовл, 8-9 хор, 10 отл.\nАвторы: Гайнанов Д.И., Гарифуллин К.Р.");
  test.setIsQuiz(true); // Включаем режим теста

  // Регистрационные поля
  test.addTextItem().setTitle("ФИО тестируемого").setRequired(true);
  test.addTextItem().setTitle("E-mail тестируемого").setRequired(true);
  test.addDateItem().setTitle("Дата тестирования").setRequired(true);

  // 10 вопросов (ровно 2 открытых вопроса, остальные с выбором ответа)
  var q1 = test.addMultipleChoiceItem();
  q1.setTitle("1. Какая физическая величина измеряется в Джоулях (Дж)?")
    .setChoices([
      q1.createChoice("Мощность", false),
      q1.createChoice("Сила", false),
      q1.createChoice("Энергия и работа", true),
      q1.createChoice("Давление", false)
    ]).setPoints(1).setRequired(true);

  var q2 = test.addMultipleChoiceItem();
  q2.setTitle("2. Первый закон Ньютона утверждает, что тело сохраняет состояние покоя или равномерного прямолинейного движения, если...")
    .setChoices([
      q2.createChoice("На него действует постоянная сила", false),
      q2.createChoice("Равнодействующая всех сил равна нулю", true),
      q2.createChoice("Скорость света постоянна", false),
      q2.createChoice("Масса тела бесконечно мала", false)
    ]).setPoints(1).setRequired(true);

  var q3 = test.addMultipleChoiceItem();
  q3.setTitle("3. Закон Ома для однородного участка электрической цепи выражается формулой:")
    .setChoices([
      q3.createChoice("I = U / R", true),
      q3.createChoice("I = U * R", false),
      q3.createChoice("P = U * I", false),
      q3.createChoice("U = I^2 * R", false)
    ]).setPoints(1).setRequired(true);

  var q4 = test.addMultipleChoiceItem();
  q4.setTitle("4. Термодинамический процесс, происходящий при постоянном объеме газа (V = const), называется:")
    .setChoices([
      q4.createChoice("Изотермический", false),
      q4.createChoice("Изобарный", false),
      q4.createChoice("Изохорный", true),
      q4.createChoice("Адиабатный", false)
    ]).setPoints(1).setRequired(true);

  var q5 = test.addMultipleChoiceItem();
  q5.setTitle("5. Какое явление волновой оптики лежит в основе передачи световых сигналов в оптическом волокне?")
    .setChoices([
      q5.createChoice("Дифракция света", false),
      q5.createChoice("Полное внутреннее отражение", true),
      q5.createChoice("Поляризация света", false),
      q5.createChoice("Дисперсия света", false)
    ]).setPoints(1).setRequired(true);

  // Открытый вопрос № 1
  var q6 = test.addTextItem();
  q6.setTitle("6. Как называется физическая векторная величина, равная произведению массы тела на его скорость (p = m*v)? (Введите одно слово)")
    .setPoints(1)
    .setRequired(true);

  var q7 = test.addMultipleChoiceItem();
  q7.setTitle("7. Выталкивающая сила (сила Архимеда), действующая на погруженное в жидкость тело, равна...")
    .setChoices([
      q7.createChoice("Весу вытесненной жидкости", true),
      q7.createChoice("Массе самого тела", false),
      q7.createChoice("Объему емкости с жидкостью", false),
      q7.createChoice("Плотности погруженного тела", false)
    ]).setPoints(1).setRequired(true);

  var q8 = test.addMultipleChoiceItem();
  q8.setTitle("8. Какой электрический заряд имеет элементарная частица нейтрон?")
    .setChoices([
      q8.createChoice("Положительный (+1)", false),
      q8.createChoice("Отрицательный (-1)", false),
      q8.createChoice("Нейтральный (ноль)", true),
      q8.createChoice("Дробный (+1/2)", false)
    ]).setPoints(1).setRequired(true);

  // Открытый вопрос № 2
  var q9 = test.addTextItem();
  q9.setTitle("9. Назовите фамилию английского ученого-физика, открывшего явление электромагнитной индукции (введите фамилию):")
    .setPoints(1)
    .setRequired(true);

  var q10 = test.addMultipleChoiceItem();
  q10.setTitle("10. Скорость распространения электромагнитного излучения (света) в вакууме приблизительно равна:")
    .setChoices([
      q10.createChoice("300 000 км/с", true),
      q10.createChoice("300 000 м/с", false),
      q10.createChoice("150 000 км/с", false),
      q10.createChoice("30 000 км/с", false)
    ]).setPoints(1).setRequired(true);

  Logger.log("--> Тест создан!");
  Logger.log("Ссылка для прохождения теста: " + test.getPublishedUrl());
  Logger.log("Ссылка для редактирования теста: " + test.getEditUrl());

  Logger.log("=== ВСЕ ФОРМЫ УСПЕШНО СОЗДАНЫ В ВАШЕМ GOOGLE ДИСКЕ! ===");
}
