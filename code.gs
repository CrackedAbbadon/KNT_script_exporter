/**
 * ============================================================
 * РАСПИСАНИЕ → GOOGLE CALENDAR
 * ============================================================
 *
 * Google Sheets → Google Calendar
 *
 * Курс и группа выбираются отдельно.
 *
 * Курс:
 *   определяет лист расписания.
 *
 * Группа:
 *   ищется на выбранном листе.
 *
 * ============================================================
 */

const CONFIG = {

  TIME_ZONE:
  Session.getScriptTimeZone() ||
  'Europe/Berlin',

  /*
   * Листы курсов.
   */
  COURSE_SHEETS: [
    '1 курс',
    '2 курс',
    '3 курс',
    '4 курс',
    '5 курс',
    '6 курс'
  ],

  /*
   * Теги Google Calendar.
   *
   * Они нужны для того, чтобы скрипт понимал,
   * какие события были созданы именно им.
   */
  TAG_SYNC:
  'schedule_sync',

  TAG_KEY:
  'schedule_key',

  TAG_GROUP:
  'schedule_group',

  TAG_COURSE:
  'schedule_course'
};


/* ============================================================
 * МЕНЮ
 * ============================================================
 */

function onOpen() {

  SpreadsheetApp
  .getUi()
  .createMenu('Расписание')

  .addItem(
    'Синхронизировать с Google Calendar',
    'showGroupDialog'
  )

  .addItem(
    'Удалить импортированные события группы',
    'showDeleteDialog'
  )

  .addSeparator()

  .addItem(
    'Показать найденные курсы',
    'showCourses'
  )

  .addToUi();
}


/* ============================================================
 * ДИАЛОГ СИНХРОНИЗАЦИИ
 * ============================================================
 */

function showGroupDialog() {

  const template =
  HtmlService.createTemplateFromFile(
    'GroupDialog'
  );

  template.mode =
  'sync';


  SpreadsheetApp
  .getUi()
  .showModalDialog(

    template
    .evaluate()
    .setWidth(500)
    .setHeight(460),

                   'Синхронизация расписания'

  );
}


/* ============================================================
 * ДИАЛОГ УДАЛЕНИЯ
 * ============================================================
 */

function showDeleteDialog() {

  const template =
  HtmlService.createTemplateFromFile(
    'GroupDialog'
  );

  template.mode =
  'delete';


  SpreadsheetApp
  .getUi()
  .showModalDialog(

    template
    .evaluate()
    .setWidth(500)
    .setHeight(460),

                   'Удаление расписания'

  );
}


/* ============================================================
 * ПОЛУЧИТЬ КУРСЫ
 * ============================================================
 */

function getCourses() {

  const ss =
  SpreadsheetApp
  .getActiveSpreadsheet();


  const result = [];


  CONFIG.COURSE_SHEETS.forEach(
    sheetName => {

      const sheet =
      ss.getSheetByName(
        sheetName
      );


      if (sheet) {

        result.push({

          name:
          sheetName

        });

      }

    }
  );


  return result;
}


/* ============================================================
 * ПОКАЗАТЬ КУРСЫ
 * ============================================================
 */

function showCourses() {

  const courses =
  getCourses();


  if (!courses.length) {

    SpreadsheetApp
    .getUi()
    .alert(
      'Листы курсов не найдены.'
    );

    return;
  }


  SpreadsheetApp
  .getUi()
  .alert(

    'Найденные курсы:\n\n' +

  courses
  .map(
    x => x.name
  )
  .join('\n')

  );
}


/* ============================================================
 * ПОЛУЧИТЬ ГРУППЫ
 * ============================================================
 */

function getGroupsForCourse(
  courseName
) {

  const sheet =
  getCourseSheet_(
    courseName
  );


  const values =
  sheet
  .getDataRange()
  .getDisplayValues();


  const groups = {};


  for (
    let r = 0;
  r < values.length;
  r++
  ) {

    for (
      let c = 0;
    c < values[r].length;
    c++
    ) {

      const value =
      normalize_(
        values[r][c]
      );


      /*
       * Прямое совпадение:
       *
       * КНТ-5
       */
      if (
        /^КНТ-\d+$/i.test(
          value
        )
      ) {

        groups[
          value.toUpperCase()
        ] = true;

      }


      /*
       * Если написано:
       *
       * 26КНТ-5
       *
       * или что-то похожее.
       */
      const match =
      value.match(
        /(?:\d+)?(КНТ-\d+)/i
      );


      if (match) {

        groups[
          match[1].toUpperCase()
        ] = true;

      }

    }

  }


  return Object
  .keys(groups)
  .sort(
    compareGroups_
  );
}


/* ============================================================
 * СОРТИРОВКА ГРУПП
 * ============================================================
 */

function compareGroups_(
  a,
  b
) {

  const na =
  Number(
    a.match(/\d+$/)?.[0] || 0
  );


  const nb =
  Number(
    b.match(/\d+$/)?.[0] || 0
  );


  return na - nb;
}


/* ============================================================
 * КАЛЕНДАРИ
 * ============================================================
 */

function getCalendars() {

  return CalendarApp
  .getAllCalendars()
  .map(
    calendar => ({

      id:
      calendar.getId(),

                 name:
                 calendar.getName(),

                 selected:
                 calendar.isSelected()

    })
  );
}


/* ============================================================
 * ОСНОВНАЯ СИНХРОНИЗАЦИЯ
 * ============================================================
 */

function syncSchedule(
  courseName,
  group,
  calendarId
) {

  if (!courseName) {

    throw new Error(
      'Не выбран курс.'
    );

  }


  if (!group) {

    throw new Error(
      'Не выбрана группа.'
    );

  }


  if (!calendarId) {

    throw new Error(
      'Не выбран календарь.'
    );

  }


  /*
   * Получаем лист курса.
   */
  const sheet =
  getCourseSheet_(
    courseName
  );


  /*
   * Получаем календарь.
   */
  const calendar =
  CalendarApp
  .getCalendarById(
    calendarId
  );


  if (!calendar) {

    throw new Error(
      'Не удалось получить доступ к выбранному календарю.'
    );

  }


  /*
   * Ищем группу.
   */
  const found =
  findGroupOnSheet_(
    sheet,
    group
  );


  if (!found) {

    throw new Error(

      'Группа ' +
    group +
    ' не найдена на листе "' +
    courseName +
    '".'

    );

  }


  /*
   * Автоматически определяем
   * период расписания.
   */
  const period =
  detectPeriod_(
    sheet
  );


  /*
   * Читаем расписание.
   */
  const events =
  parseGroup_(
    sheet,
    found.headerRow,
    found.groupCol,
    group,
    period
  );


  /*
   * Синхронизируем.
   */
  const stats =
  syncEvents_(
    calendar,
    events,
    group,
    courseName,
    period
  );


  return {

    course:
    courseName,

    group:
    group,

    sheet:
    sheet.getName(),

    period:
    formatPeriod_(
      period
    ),

    total:
    events.length,

    created:
    stats.created,

    updated:
    stats.updated,

    deleted:
    stats.deleted

  };
}


/* ============================================================
 * УДАЛЕНИЕ СОБЫТИЙ ГРУППЫ
 * ============================================================
 */

function deleteGroupEvents(
  courseName,
  group,
  calendarId
) {

  if (!courseName) {

    throw new Error(
      'Не выбран курс.'
    );

  }


  if (!group) {

    throw new Error(
      'Не выбрана группа.'
    );

  }


  if (!calendarId) {

    throw new Error(
      'Не выбран календарь.'
    );

  }


  const calendar =
  CalendarApp
  .getCalendarById(
    calendarId
  );


  if (!calendar) {

    throw new Error(
      'Не удалось получить доступ к выбранному календарю.'
    );

  }


  const start =
  new Date(
    2020,
    0,
    1
  );


  const end =
  new Date(
    2035,
    11,
    31
  );


  const events =
  calendar.getEvents(
    start,
    end
  );


  let deleted =
  0;


  events.forEach(
    event => {

      if (
        event.getTag(
          CONFIG.TAG_SYNC
        ) !== '1'
      ) {
        return;
      }


      if (
        event.getTag(
          CONFIG.TAG_GROUP
        ) !== group
      ) {
        return;
      }


      if (
        event.getTag(
          CONFIG.TAG_COURSE
        ) !== courseName
      ) {
        return;
      }


      event.deleteEvent();

      deleted++;

    }
  );


  return {

    course:
    courseName,

    group:
    group,

    deleted:
    deleted

  };
}


/* ============================================================
 * ПОЛУЧИТЬ ЛИСТ КУРСА
 * ============================================================
 */

function getCourseSheet_(
  courseName
) {

  const ss =
  SpreadsheetApp
  .getActiveSpreadsheet();


  const sheet =
  ss.getSheetByName(
    courseName
  );


  if (!sheet) {

    throw new Error(

      'Лист "' +
    courseName +
    '" не найден.'

    );

  }


  return sheet;
}


/* ============================================================
 * ПОИСК ГРУППЫ
 * ============================================================
 */

function findGroupOnSheet_(
  sheet,
  group
) {

  const values =
  sheet
  .getDataRange()
  .getDisplayValues();


  const wanted =
  normalizeGroup_(
    group
  );


  for (
    let r = 0;
  r < values.length;
  r++
  ) {

    for (
      let c = 0;
    c < values[r].length;
    c++
    ) {

      const value =
      normalize_(
        values[r][c]
      );


      const normalized =
      normalizeGroup_(
        value
      );


      if (
        normalized === wanted
      ) {

        return {

          sheet:
          sheet,

          headerRow:
          r + 1,

          groupCol:
          c + 1

        };

      }

    }

  }


  return null;
}


/* ============================================================
 * НОРМАЛИЗАЦИЯ ГРУППЫ
 * ============================================================
 */

function normalizeGroup_(
  value
) {

  const text =
  normalize_(
    value
  )
  .toUpperCase();


  const match =
  text.match(
    /(?:\d+)?(КНТ-\d+)/
  );


  if (!match) {

    return text;

  }


  return match[1];
}


/* ============================================================
 * АВТОМАТИЧЕСКОЕ ОПРЕДЕЛЕНИЕ ПЕРИОДА
 * ============================================================
 */

function detectPeriod_(
  sheet
) {

  const values =
  sheet
  .getDataRange()
  .getDisplayValues();


  const year =
  guessYear_(
    values
  );


  const dates = [];


  for (
    let r = 0;
  r < values.length;
  r++
  ) {

    for (
      let c = 0;
    c < values[r].length;
    c++
    ) {

      const text =
      String(
        values[r][c] || ''
      );


      parseAllDatesFromText_(
        text,
        year
      ).forEach(
        date => {

          dates.push(
            date
          );

        }
      );


      findRanges_(
        text,
        year
      ).forEach(
        range => {

          dates.push(
            range.start
          );

          dates.push(
            range.end
          );

        }
      );

    }

  }


  if (!dates.length) {

    throw new Error(

      'На листе "' +
    sheet.getName() +
    '" не удалось найти даты расписания.'

    );

  }


  dates.sort(
    (a, b) =>
    a.getTime() -
    b.getTime()
  );


  return {

    start:
    startOfDay_(
      dates[0]
    ),

    end:
    endOfDay_(
      dates[
        dates.length - 1
      ]
    )

  };
}


/* ============================================================
 * ОПРЕДЕЛЕНИЕ ГОДА
 * ============================================================
 */

function guessYear_(
  values
) {

  for (
    let r = 0;
  r < values.length;
  r++
  ) {

    for (
      let c = 0;
    c < values[r].length;
    c++
    ) {

      const text =
      String(
        values[r][c] || ''
      );


      const match =
      text.match(
        /\b\d{1,2}\.\d{1,2}\.(\d{2,4})\b/
      );


      if (match) {

        let year =
        Number(
          match[1]
        );


        if (
          year < 100
        ) {

          year += 2000;

        }


        return year;

      }

    }

  }


  return new Date()
  .getFullYear();
}


/* ============================================================
 * ПОЛУЧИТЬ ВСЕ ДАТЫ
 * ============================================================
 */

function parseAllDatesFromText_(
  text,
  defaultYear
) {

  const regex =
  /(?<!\d)(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?(?!\d)/g;


  const result = [];


  let match;


  while (
    (match =
    regex.exec(
      text
    )) !== null
  ) {

    const day =
    Number(
      match[1]
    );


    const month =
    Number(
      match[2]
    );


    let year =
    defaultYear;


    if (
      match[3]
    ) {

      year =
      Number(
        match[3]
      );


      if (
        year < 100
      ) {

        year += 2000;

      }

    }


    const date =
    new Date(
      year,
      month - 1,
      day
    );


    if (

      date.getFullYear() === year &&

      date.getMonth() ===
      month - 1 &&

      date.getDate() ===
      day

    ) {

      result.push(
        date
      );

    }

  }


  return result;
}


/* ============================================================
 * ДИАПАЗОНЫ ДАТ
 * ============================================================
 */

function findRanges_(
  text,
  year
) {

  const regex =
  /(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\s*[-–—]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?/g;


  const result = [];


  let match;


  while (
    (match =
    regex.exec(
      text
    )) !== null
  ) {

    let startYear =
    year;


    let endYear =
    year;


    if (
      match[3]
    ) {

      startYear =
      Number(
        match[3]
      );


      if (
        startYear < 100
      ) {

        startYear += 2000;

      }

    }


    if (
      match[6]
    ) {

      endYear =
      Number(
        match[6]
      );


      if (
        endYear < 100
      ) {

        endYear += 2000;

      }

    }


    result.push({

      start:
      new Date(

        startYear,

        Number(match[2]) - 1,

               Number(match[1])

      ),

      end:
      new Date(

        endYear,

        Number(match[5]) - 1,

               Number(match[4])

      )

    });

  }


  return result;
}


/* ============================================================
 * ФОРМАТ ПЕРИОДА
 * ============================================================
 */

function formatPeriod_(
  period
) {

  return (

    Utilities.formatDate(

      period.start,

      CONFIG.TIME_ZONE,

      'dd.MM.yyyy'

    ) +

    ' — ' +

  Utilities.formatDate(

    period.end,

    CONFIG.TIME_ZONE,

    'dd.MM.yyyy'

  )

  );
}


/* ============================================================
 * ДНИ НЕДЕЛИ
 * ============================================================
 */

const DAY_MAP_ = {

  'вс': 0,
  'воскресенье': 0,

  'пн': 1,
  'понедельник': 1,

  'вт': 2,
  'вторник': 2,

  'ср': 3,
  'среда': 3,

  'чт': 4,
  'четверг': 4,

  'пт': 5,
  'пятница': 5,

  'сб': 6,
  'суббота': 6

};


/* ============================================================
 * ПАРСИНГ ГРУППЫ
 * ============================================================
 */

function parseGroup_(
  sheet,
  headerRow,
  groupCol,
  group,
  period
) {

  const values =
  sheet
  .getDataRange()
  .getDisplayValues();


  const events = [];


  let currentDay =
  null;


  for (
    let r = headerRow;
  r < values.length;
  r++
  ) {

    const rowNumber =
    r + 1;


    /*
     * День недели.
     */
    const dayValue =
    normalize_(
      values[r][1] || ''
    );


    const dayLower =
    dayValue.toLowerCase();


    if (
      Object.prototype
      .hasOwnProperty
      .call(
        DAY_MAP_,
        dayLower
      )
    ) {

      currentDay =
      DAY_MAP_[
        dayLower
      ];

    }


    /*
     * Время.
     */
    const timeValue =
    normalize_(
      values[r][2] || ''
    );


    const time =
    parseTime_(
      timeValue
    );


    if (

      currentDay === null ||
      !time

    ) {

      continue;

    }


    /*
     * Занятие.
     */
    const subjectCell =
    values[r][groupCol - 1] ||
    '';


    if (
      !normalize_(
        subjectCell
      )
    ) {

      continue;

    }


    /*
     * Аудитория.
     */
    const roomCell =
    values[r][groupCol] ||
    '';


    /*
     * Корпус.
     */
    const buildingCell =
    values[r][groupCol + 1] ||
    '';


    /*
     * Разделяем занятия,
     * если в ячейке их несколько.
     */
    const blocks =
    splitBlocks_(
      subjectCell
    );


    const rooms =
    splitMeta_(
      roomCell,
      blocks.length
    );


    const buildings =
    splitMeta_(
      buildingCell,
      blocks.length
    );


    blocks.forEach(
      (block, index) => {

        const parsed =
        parseBlock_(
          block,
          period
        );


        let dates = [];


        /*
         * Конкретные даты.
         */
        if (
          parsed.mode ===
          'explicit'
        ) {

          dates =
          parsed.dates;

        }


        /*
         * С определённой даты
         * каждую неделю.
         */
        else if (
          parsed.mode ===
          'weekly_from'
        ) {

          dates =
          weeklyDates_(
            parsed.startDate,
            period.end,
            currentDay
          );

        }


        /*
         * Диапазон.
         */
        else if (
          parsed.mode ===
          'range'
        ) {

          dates =
          weeklyDates_(
            parsed.rangeStart,
            parsed.rangeEnd,
            currentDay
          );

        }


        /*
         * Обычное еженедельное занятие.
         */
        else {

          dates =
          weeklyDates_(
            period.start,
            period.end,
            currentDay
          );

        }


        dates.forEach(
          dateValue => {

            if (

              dateValue <
              period.start ||

              dateValue >
              period.end

            ) {

              return;

            }


            const room =
            normalize_(
              rooms[index] || ''
            );


            const building =
            normalize_(
              buildings[index] || ''
            );


            events.push({

              key:
              makeEventKey_(
                group,
                rowNumber,
                index,
                dateValue
              ),

              group:
              group,

              date:
              dateValue,

              startHour:
              time.startHour,

              startMinute:
              time.startMinute,

              endHour:
              time.endHour,

              endMinute:
              time.endMinute,

              subject:
              parsed.subject ||
              'Занятие',

              teacher:
              parsed.teacher ||
              '',

              room:
              room,

              building:
              building,

              sheet:
              sheet.getName(),

                        row:
                        rowNumber,

                        block:
                        index

            });

          }
        );

      }
    );

  }


  /*
   * Удаляем дубликаты.
   */
  const unique = {};


  events.forEach(
    event => {

      unique[
        event.key
      ] = event;

    }
  );


  return Object
  .keys(unique)
  .map(
    key =>
    unique[key]
  )
  .sort(
    (a, b) => {

      const dateDiff =
      a.date.getTime() -
      b.date.getTime();


      if (
        dateDiff !== 0
      ) {

        return dateDiff;

      }


      return (

        a.startHour * 60 +
        a.startMinute

      ) -

      (

        b.startHour * 60 +
        b.startMinute

      );

    }
  );
}


/* ============================================================
 * РАЗБОР ЗАНЯТИЯ
 * ============================================================
 */

function parseBlock_(
  block,
  period
) {

  let text =
  String(
    block || ''
  )
  .replace(
    /\r\n/g,
    '\n'
  )
  .replace(
    /\r/g,
    '\n'
  )
  .trim();


  /*
   * ==========================================================
   * ДИАПАЗОН
   *
   * 07.09.-19.10.
   * ==========================================================
   */

  let match =
  text.match(

    /^\s*,?\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\s*[-–—]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?\s*[,:;]?\s*/i

  );


  if (match) {

    let startYear =
    period.start.getFullYear();


    let endYear =
    period.start.getFullYear();


    if (
      match[3]
    ) {

      startYear =
      Number(
        match[3]
      );


      if (
        startYear < 100
      ) {

        startYear += 2000;

      }

    }


    if (
      match[6]
    ) {

      endYear =
      Number(
        match[6]
      );


      if (
        endYear < 100
      ) {

        endYear += 2000;

      }

    }


    const startDate =
    new Date(

      startYear,

      Number(match[2]) - 1,

             Number(match[1])

    );


    const endDate =
    new Date(

      endYear,

      Number(match[5]) - 1,

             Number(match[4])

    );


    text =
    text
    .substring(
      match[0].length
    )
    .trim();


    text =
    removeDateListPrefix_(
      text
    );


    const cleaned =
    cleanSubjectTeacher_(
      text
    );


    return {

      mode:
      'range',

      rangeStart:
      startDate,

      rangeEnd:
      endDate,

      subject:
      cleaned.subject,

      teacher:
      cleaned.teacher

    };
  }


  /*
   * ==========================================================
   * С ДАТЫ
   *
   * с 09.09.
   * ==========================================================
   */

  match =
  text.match(

    /^\s*,?\s*с\s+(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?\s*[,:;]?\s*/i

  );


  if (match) {

    let year =
    period.start.getFullYear();


    if (
      match[3]
    ) {

      year =
      Number(
        match[3]
      );


      if (
        year < 100
      ) {

        year += 2000;

      }

    }


    const startDate =
    new Date(

      year,

      Number(match[2]) - 1,

             Number(match[1])

    );


    text =
    text
    .substring(
      match[0].length
    )
    .trim();


    text =
    removeDateListPrefix_(
      text
    );


    const cleaned =
    cleanSubjectTeacher_(
      text
    );


    return {

      mode:
      'weekly_from',

      startDate:
      startDate,

      subject:
      cleaned.subject,

      teacher:
      cleaned.teacher

    };
  }


  /*
   * ==========================================================
   * СПИСОК КОНКРЕТНЫХ ДАТ
   *
   * 04.09., 18.09., 02.10., 16.10.
   *
   * или
   *
   * , 18.09., 02.10., 16.10.
   * ==========================================================
   */

  const explicit =
  extractDateListPrefix_(
    text,
    period.start.getFullYear()
  );


  if (

    explicit &&
    explicit.dates.length

  ) {

    text =
    explicit.rest;


    const cleaned =
    cleanSubjectTeacher_(
      text
    );


    return {

      mode:
      'explicit',

      dates:
      explicit.dates,

      subject:
      cleaned.subject,

      teacher:
      cleaned.teacher

    };
  }


  /*
   * ==========================================================
   * ОБЫЧНОЕ ЗАНЯТИЕ
   * ==========================================================
   */

  const cleaned =
  cleanSubjectTeacher_(
    text
  );


  return {

    mode:
    'none',

    subject:
    cleaned.subject,

    teacher:
    cleaned.teacher

  };
}


/* ============================================================
 * ИЗВЛЕЧЕНИЕ СПИСКА ДАТ
 * ============================================================
 */

function extractDateListPrefix_(
  text,
  defaultYear
) {

  let rest =
  String(
    text || ''
  ).trim();


  const dates = [];


  /*
   * Убираем начальные запятые.
   */
  rest =
  rest.replace(
    /^[,\s;]+/,
    ''
  );


  while (true) {

    const match =
    rest.match(

      /^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?\s*(?:[,;]\s*)?/

    );


    if (!match) {
      break;
    }


    const day =
    Number(
      match[1]
    );


    const month =
    Number(
      match[2]
    );


    let year =
    defaultYear;


    if (
      match[3]
    ) {

      year =
      Number(
        match[3]
      );


      if (
        year < 100
      ) {

        year += 2000;

      }

    }


    const date =
    new Date(

      year,

      month - 1,

      day

    );


    if (

      date.getFullYear() !== year ||

      date.getMonth() !==
      month - 1 ||

      date.getDate() !==
      day

    ) {

      break;

    }


    dates.push(
      date
    );


    rest =
    rest
    .substring(
      match[0].length
    )
    .trim();


    rest =
    rest.replace(
      /^[,;]\s*/,
      ''
    );

  }


  if (!dates.length) {

    return null;

  }


  return {

    dates:
    dates,

    rest:
    rest

  };
}


/* ============================================================
 * УДАЛИТЬ ОСТАВШИЕСЯ ДАТЫ
 * ============================================================
 */

function removeDateListPrefix_(
  text
) {

  let value =
  String(
    text || ''
  )
  .trim();


  while (true) {

    const match =
    value.match(

      /^\s*[,;]\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?\s*/

    );


    if (!match) {
      break;
    }


    value =
    value
    .substring(
      match[0].length
    )
    .trim();

  }


  const extracted =
  extractDateListPrefix_(
    value,
    new Date().getFullYear()
  );


  if (

    extracted &&
    extracted.dates.length

  ) {

    return extracted.rest;

  }


  return value;
}

/* ============================================================
 * НАЗВАНИЕ + ПРЕПОДАВАТЕЛЬ
 * ============================================================
 *
 * Поддерживает варианты:
 *
 * Математический анализ - семинар
 * Чистякова С.А.
 *
 * Математический анализ - семинар - Чистякова С.А.
 *
 * Программирование C/C++ - лекция - Лупанова Е.А.
 *
 * Технологии программирования - Улитин И.Б.
 *
 * Основы российской государственности - семинар -
 * 26_27_М_ОРГ_Г_1115950_14 - Константинова Т.Н.
 *
 * ============================================================
 */

function cleanSubjectTeacher_(text) {

  let value =
  String(text || '')
  .replace(/\r\n/g, '\n')
  .replace(/\r/g, '\n');

  /*
   * Переносы строк внутри одной записи
   * превращаем в пробелы.
   */
  value =
  value
  .replace(/\n+/g, ' ')
  .replace(/\s+/g, ' ')
  .trim();


  /*
   * Убираем мусор в начале.
   */
  value =
  value.replace(
    /^[,\s.;:]+/,
    ''
  );


  /*
   * Убираем мусор в конце.
   */
  value =
  value.replace(
    /[,\s.;:]+$/,
    ''
  );


  /*
   * Если даты почему-то всё ещё остались
   * в начале строки — удаляем их.
   */
  value =
  value.replace(

    /^(?:,\s*)?(?:\d{1,2}\.\d{1,2}(?:\.\d{2,4})?\s*[,;.]?\s*)+/,

                ''

  ).trim();


  /*
   * ==========================================================
   * ПРЕПОДАВАТЕЛЬ В КОНЦЕ СТРОКИ
   * ==========================================================
   *
   * Важно:
   *
   * преподаватель может быть отделён:
   *
   * "- Чистякова С.А."
   *
   * или просто пробелом:
   *
   * "семинар Чистякова С.А."
   *
   * или переносом строки, который выше
   * уже превратился в пробел.
   *
   * ==========================================================
   */


  /*
   * Фамилия + инициалы:
   *
   * Чистякова С.А.
   * Чистякова С. А.
   * Чистякова С.А
   */
  const shortTeacherRegex =

  /([А-ЯЁA-Z][а-яёa-z-]+)\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?\s*$/;


  /*
   * Фамилия Имя + инициалы:
   *
   * Иванов Иван И.И.
   */
  const longTeacherRegex =

  /([А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z][а-яёa-z-]+)\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?\s*$/;


  let teacherMatch =
  value.match(
    longTeacherRegex
  );


  if (!teacherMatch) {

    teacherMatch =
    value.match(
      shortTeacherRegex
    );

  }


  if (teacherMatch) {

    const teacher =
    normalizeTeacher_(
      teacherMatch[0]
    );


    let before =
    value.substring(
      0,
      teacherMatch.index
    ).trim();


    /*
     * Убираем разделитель перед преподавателем:
     *
     * -
     * –
     * —
     *
     * Например:
     *
     * "Математический анализ - семинар -"
     *
     * превращается в:
     *
     * "Математический анализ - семинар"
     */
    before =
    before.replace(
      /\s*[-–—]\s*$/,
      ''
    ).trim();


    /*
     * Если преподаватель идёт сразу после
     * названия/типа занятия без дефиса:
     *
     * "Математический анализ - семинар Чистякова С.А."
     *
     * before уже будет:
     *
     * "Математический анализ - семинар"
     *
     * поэтому ничего дополнительно делать не нужно.
     */


    if (before) {

      return {

        subject:
        cleanSubjectText_(
          before
        ),

        teacher:
        teacher

      };

    }

  }


  /*
   * ==========================================================
   * ДОПОЛНИТЕЛЬНЫЙ ВАРИАНТ
   * ==========================================================
   *
   * Иногда преподаватель может быть отделён
   * несколькими дефисами:
   *
   * Предмет - семинар - код - Преподаватель
   *
   * Здесь используем разделение по " - ".
   */

  const parts =
  value
  .split(
    /\s*[-–—]\s*/
  )
  .map(
    x => x.trim()
  )
  .filter(
    Boolean
  );


  if (
    parts.length >= 2
  ) {

    const candidate =
    parts[
      parts.length - 1
    ];


    if (
      isTeacherName_(
        candidate
      )
    ) {

      return {

        subject:
        cleanSubjectText_(
          parts
          .slice(
            0,
            -1
          )
          .join(
            ' - '
          )
        ),

        teacher:
        normalizeTeacher_(
          candidate
        )

      };

    }

  }


  /*
   * Преподаватель не найден.
   */
  return {

    subject:
    cleanSubjectText_(
      value
    ),

    teacher:
    ''

  };
}


/* ============================================================
 * ОЧИСТКА НАЗВАНИЯ ПРЕДМЕТА
 * ============================================================
 */

function cleanSubjectText_(
  text
) {

  return String(
    text || ''
  )

  .replace(
    /\s+/g,
    ' '
  )

  .replace(
    /\s*[-–—]\s*$/,
    ''
  )

  .replace(
    /^[,\s.;:]+/,
    ''
  )

  .replace(
    /[,\s.;:]+$/,
    ''
  )

  .trim();
}


/* ============================================================
 * ПРОВЕРКА ПРЕПОДАВАТЕЛЯ
 * ============================================================
 */

function isTeacherName_(
  value
) {

  const text =
  String(
    value || ''
  )
  .replace(
    /\s+/g,
    ' '
  )
  .trim();


  /*
   * Фамилия И.О.
   *
   * Чистякова С.А.
   * Чистякова С. А.
   */
  if (

    /^[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?$/
    .test(
      text
    )

  ) {

    return true;

  }


  /*
   * Фамилия Имя И.О.
   *
   * Иванов Иван И.И.
   */
  if (

    /^[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?$/
    .test(
      text
    )

  ) {

    return true;

  }


  return false;
}


/* ============================================================
 * НОРМАЛИЗАЦИЯ ПРЕПОДАВАТЕЛЯ
 * ============================================================
 */

function normalizeTeacher_(
  teacher
) {

  if (!teacher) {

    return '';

  }


  return String(
    teacher
  )

  .replace(
    /^Преподаватель\s*:\s*/i,
    ''
  )

  .replace(
    /\s+/g,
    ' '
  )

  .trim()

  .replace(
    /[,\s]+$/,
    ''
  );
}

/* ============================================================
 * ПРОВЕРКА ПРЕПОДАВАТЕЛЯ
 * ============================================================
 */

function isTeacherName_(
  value
) {

  const text =
  String(
    value || ''
  )
  .trim();


  /*
   * Фамилия И.О.
   */
  if (

    /^[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?$/
    .test(
      text
    )

  ) {

    return true;

  }


  /*
   * Фамилия Имя И.О.
   */
  if (

    /^[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z][а-яёa-z-]+\s+[А-ЯЁA-Z]\.?\s*[А-ЯЁA-Z]\.?$/
    .test(
      text
    )

  ) {

    return true;

  }


  return false;
}


/* ============================================================
 * НОРМАЛИЗАЦИЯ ФИО
 * ============================================================
 */

function normalizeTeacher_(teacher) {
  if (!teacher) {
    return '';
  }

  teacher = String(teacher)
  .replace(/\s+/g, ' ')
  .trim();

  // Убираем служебные подписи
  teacher = teacher
  .replace(/^Преподаватель\s*:\s*/i, '')
  .trim();

  return teacher;
}
/* ============================================================
 * РАЗДЕЛЕНИЕ ЗАНЯТИЙ
 * ============================================================
 */

function splitBlocks_(
  value
) {

  const text =
  String(
    value || ''
  )
  .replace(
    /\r\n/g,
    '\n'
  )
  .replace(
    /\r/g,
    '\n'
  )
  .trim();


  if (!text) {

    return [];

  }


  /*
   * Занятия в одной ячейке
   * разделены пустой строкой.
   */
  return text
  .split(
    /\n\s*\n+/
  )
  .map(
    x => x.trim()
  )
  .filter(
    Boolean
  );
}


/* ============================================================
 * РАЗДЕЛЕНИЕ АУДИТОРИЙ / КОРПУСОВ
 * ============================================================
 */

function splitMeta_(
  value,
  count
) {

  const text =
  String(
    value || ''
  );


  if (!text.trim()) {

    return new Array(
      count
    ).fill('');

  }


  const parts =
  text
  .split(
    /\n+/
  )
  .map(
    x => x.trim()
  )
  .filter(
    Boolean
  );


  if (
    parts.length === count
  ) {

    return parts;

  }


  if (
    parts.length === 1
  ) {

    return new Array(
      count
    ).fill(
      parts[0]
    );

  }


  const result = [];


  for (
    let i = 0;
  i < count;
  i++
  ) {

    result.push(

      parts[i] !== undefined

      ? parts[i]

      : parts[
        parts.length - 1
      ]

    );

  }


  return result;
}


/* ============================================================
 * ВРЕМЯ
 * ============================================================
 */

function parseTime_(
  value
) {

  const match =
  String(
    value || ''
  ).match(

    /(\d{1,2}):(\d{2})\s*[-–—]\s*(\d{1,2}):(\d{2})/

  );


  if (!match) {

    return null;

  }


  return {

    startHour:
    Number(
      match[1]
    ),

    startMinute:
    Number(
      match[2]
    ),

    endHour:
    Number(
      match[3]
    ),

    endMinute:
    Number(
      match[4]
    )

  };
}


/* ============================================================
 * ДАТЫ ПО НЕДЕЛЯМ
 * ============================================================
 */

function weeklyDates_(
  start,
  end,
  weekday
) {

  const result = [];


  let current =
  startOfDay_(
    new Date(
      start
    )
  );


  const finalDate =
  endOfDay_(
    new Date(
      end
    )
  );


  while (

    current.getDay() !== weekday &&
    current <= finalDate

  ) {

    current.setDate(
      current.getDate() + 1
    );

  }


  while (
    current <= finalDate
  ) {

    result.push(
      new Date(
        current
      )
    );


    current.setDate(
      current.getDate() + 7
    );

  }


  return result;
}


/* ============================================================
 * СИНХРОНИЗАЦИЯ СОБЫТИЙ
 * ============================================================
 */

function syncEvents_(
  calendar,
  events,
  group,
  courseName,
  period
) {

  const managed =
  getManagedEvents_(
    calendar,
    group,
    courseName,
    period
  );


  const byKey = {};


  managed.forEach(
    item => {

      if (
        item.key
      ) {

        byKey[
          item.key
        ] =
        item.event;

      }

    }
  );


  const desired = {};


  events.forEach(
    event => {

      desired[
        event.key
      ] =
      event;

    }
  );


  let created =
  0;


  let updated =
  0;


  let deleted =
  0;


  /*
   * Создание / обновление.
   */
  events.forEach(
    event => {

      const oldEvent =
      byKey[
        event.key
      ];


      if (!oldEvent) {

        createCalendarEvent_(
          calendar,
          event,
          courseName
        );


        created++;

        return;
      }


      if (

        updateCalendarEventIfNeeded_(
          oldEvent,
          event,
          courseName
        )

      ) {

        updated++;

      }

    }
  );


  /*
   * Удаление старых событий.
   */
  managed.forEach(
    item => {

      if (

        item.key &&
        !desired[
          item.key
        ]

      ) {

        item.event
        .deleteEvent();


        deleted++;

      }

    }
  );


  return {

    created:
    created,

    updated:
    updated,

    deleted:
    deleted

  };
}


/* ============================================================
 * ПОЛУЧЕНИЕ ИМПОРТИРОВАННЫХ СОБЫТИЙ
 * ============================================================
 */

function getManagedEvents_(
  calendar,
  group,
  courseName,
  period
) {

  const events =
  calendar.getEvents(

    startOfDay_(
      period.start
    ),

    endOfDay_(
      period.end
    )

  );


  const result = [];


  events.forEach(
    event => {

      /*
       * Только наши события.
       */
      if (

        event.getTag(
          CONFIG.TAG_SYNC
        ) !== '1'

      ) {

        return;

      }


      /*
       * Группа.
       */
      if (

        event.getTag(
          CONFIG.TAG_GROUP
        ) !== group

      ) {

        return;

      }


      /*
       * Курс.
       */
      if (

        event.getTag(
          CONFIG.TAG_COURSE
        ) !== courseName

      ) {

        return;

      }


      result.push({

        event:
        event,

        key:
        event.getTag(
          CONFIG.TAG_KEY
        ) || ''

      });

    }
  );


  return result;
}


/* ============================================================
 * СОЗДАНИЕ СОБЫТИЯ
 * ============================================================
 */

function createCalendarEvent_(
  calendar,
  data,
  courseName
) {

  const start =
  makeDateTime_(
    data.date,
    data.startHour,
    data.startMinute
  );


  const end =
  makeDateTime_(
    data.date,
    data.endHour,
    data.endMinute
  );


  const event =
  calendar.createEvent(

    data.subject,

    start,

    end,

    {

      description:
      buildDescription_(
        data
      ),

      location:
      buildLocation_(
        data
      )

    }

  );


  /*
   * Техническая информация
   * находится в тегах.
   */
  setEventTags_(
    event,
    data,
    courseName
  );
}


/* ============================================================
 * ОБНОВЛЕНИЕ СОБЫТИЯ
 * ============================================================
 */

function updateCalendarEventIfNeeded_(
  event,
  data,
  courseName
) {

  const start =
  makeDateTime_(
    data.date,
    data.startHour,
    data.startMinute
  );


  const end =
  makeDateTime_(
    data.date,
    data.endHour,
    data.endMinute
  );


  const expectedDescription =
  buildDescription_(
    data
  );


  const expectedLocation =
  buildLocation_(
    data
  );


  let changed =
  false;


  /*
   * Название.
   */
  if (

    event.getTitle() !==
    data.subject

  ) {

    event.setTitle(
      data.subject
    );


    changed =
    true;

  }


  /*
   * Время.
   */
  if (

    event.getStartTime()
    .getTime() !==
    start.getTime() ||

    event.getEndTime()
    .getTime() !==
    end.getTime()

  ) {

    event.setTime(
      start,
      end
    );


    changed =
    true;

  }


  /*
   * Описание.
   */
  if (

    event.getDescription() !==
    expectedDescription

  ) {

    event.setDescription(
      expectedDescription
    );


    changed =
    true;

  }


  /*
   * Место.
   */
  if (

    event.getLocation() !==
    expectedLocation

  ) {

    event.setLocation(
      expectedLocation
    );


    changed =
    true;

  }


  /*
   * Теги.
   */
  setEventTags_(
    event,
    data,
    courseName
  );


  return changed;
}


/* ============================================================
 * ТЕГИ СОБЫТИЯ
 * ============================================================
 */

function setEventTags_(
  event,
  data,
  courseName
) {

  event.setTag(
    CONFIG.TAG_SYNC,
    '1'
  );


  event.setTag(
    CONFIG.TAG_KEY,
    data.key
  );


  event.setTag(
    CONFIG.TAG_GROUP,
    data.group
  );


  event.setTag(
    CONFIG.TAG_COURSE,
    courseName
  );
}

/* ============================================================
 * ОПИСАНИЕ СОБЫТИЯ
 * ============================================================
 *
 * Например:
 *
 * Преподаватель: Чистякова С.А.
 * Аудитория: 304
 *
 * ============================================================
 */

function buildDescription_(
  data
) {

  const parts = [];


  /*
   * Преподаватель.
   */
  if (
    data.teacher
  ) {

    parts.push(
      'Преподаватель: ' +
    data.teacher
    );

  }


  /*
   * Аудитория.
   */
  const room =
  normalize_(
    data.room
  );


  if (

    room &&
    room !== '-' &&
    room !== '—'

  ) {

    let roomText =
    room;


    /*
     * Если уже написано:
     *
     * ауд. 304
     *
     * ничего не добавляем.
     */
    if (
      !/^ауд\.?/i.test(
        roomText
      )
    ) {

      roomText =
      'ауд. ' +
      roomText;

    }


    parts.push(
      'Аудитория: ' +
    roomText
    );

  }


  return parts.join(
    '\n'
  );
}
/* ============================================================
 * МЕСТО
 * ============================================================
 */

function buildLocation_(
  data
) {

  let building =
  normalize_(
    data.building
  );


  let room =
  normalize_(
    data.room
  );


  const buildingNormalized =
  building

  .replace(
    /\./g,
    ''
  )

  .replace(
    /\s+/g,
    ''
  )

  .toLowerCase();


  /*
   * Онлайн.
   */
  if (

    buildingNormalized === 'online' ||

    buildingNormalized === 'онлайн'

  ) {

    return 'Онлайн';

  }


  /*
   * Корпуса.
   */
  const BUILDINGS = {

    'бп':
    'БП, ул. Б. Печёрская, 25/12',

    'бпок':
    'БПок, ул. Б. Покровская, 22',

    'к':
    'К, ул. Костина, 2Б',

    'л':
    'Л, ул. Львовская, 1В',

    'р':
    'Р, ул. Родионова, 136',

    'с':
    'С, Сормовское ш., 30'

  };


  let buildingAddress =
  BUILDINGS[
    buildingNormalized
  ] || building;


  /*
   * Аудитория.
   */
  let roomText = '';


  if (

    room &&
    room !== '-' &&
    room !== '—'

  ) {

    if (

      /^ауд\.?\s*/i.test(
        room
      )

    ) {

      roomText =
      room;

    } else {

      roomText =
      'ауд. ' +
      room;

    }

  }


  const parts = [];


  if (
    roomText
  ) {

    parts.push(
      roomText
    );

  }


  if (
    buildingAddress
  ) {

    parts.push(
      buildingAddress
    );

  }


  return parts.join(
    ', '
  );
}

/* ============================================================
 * KEY СОБЫТИЯ
 * ============================================================
 */

function makeEventKey_(
  group,
  row,
  block,
  dateValue
) {

  return [

    group,

    Utilities.formatDate(
      dateValue,
      CONFIG.TIME_ZONE,
      'yyyy-MM-dd'
    ),

    'ROW=' +
    row,

    'BLOCK=' +
    block

  ].join('|');
}


/* ============================================================
 * ДАТА + ВРЕМЯ
 * ============================================================
 */

function makeDateTime_(
  dateValue,
  hour,
  minute
) {

  return new Date(

    dateValue.getFullYear(),

                  dateValue.getMonth(),

                  dateValue.getDate(),

                  hour,

                  minute,

                  0,

                  0

  );
}


/* ============================================================
 * НАЧАЛО ДНЯ
 * ============================================================
 */

function startOfDay_(
  dateValue
) {

  return new Date(

    dateValue.getFullYear(),

                  dateValue.getMonth(),

                  dateValue.getDate(),

                  0,
                  0,
                  0,
                  0

  );
}


/* ============================================================
 * КОНЕЦ ДНЯ
 * ============================================================
 */

function endOfDay_(
  dateValue
) {

  return new Date(

    dateValue.getFullYear(),

                  dateValue.getMonth(),

                  dateValue.getDate(),

                  23,
                  59,
                  59,
                  999

  );
}


/* ============================================================
 * ОБЩАЯ НОРМАЛИЗАЦИЯ
 * ============================================================
 */

function normalize_(
  value
) {

  return String(
    value || ''
  )

  .replace(
    /\s+/g,
    ' '
  )

  .trim();
}
