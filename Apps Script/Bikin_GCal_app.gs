/************************************************
 * BIKIN GCAL APP
 * VERSION 3 - FIX GUEST + WEEKLY H+1 ISSUE
 ************************************************/

const SHEET_NAME = "Sheet1";

/************************************************
 * WEB APP
 ************************************************/

function doGet() {

  return HtmlService
    .createTemplateFromFile(
      "Bikin_GCal_Index"
    )
    .evaluate()
    .setTitle(
      "Event Generator"
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(
      filename
    )
    .getContent();

}

/************************************************
 * VALIDASI EMAIL
 ************************************************/

function validateEmails(emailString) {

  if (
    !emailString ||
    String(emailString).trim() === ""
  ) {

    return true;

  }

  const emails =
    String(emailString)
      .split(",")
      .map(email => email.trim())
      .filter(Boolean);

  const regex =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  const invalid =
    emails.filter(
      email => !regex.test(email)
    );

  if (
    invalid.length > 0
  ) {

    throw new Error(
      "Email tidak valid:\n\n" +
      invalid.join("\n")
    );

  }

  return true;

}

/************************************************
 * NORMALISASI EMAIL GUEST
 ************************************************/

function normalizeGuests_(emailString) {

  if (
    !emailString ||
    String(emailString).trim() === ""
  ) {

    return "";

  }

  return String(emailString)
    .split(",")
    .map(email => email.trim())
    .filter(Boolean)
    .join(",");

}

/************************************************
 * PARSE DATE
 * Input: yyyy-mm-dd
 ************************************************/

function parseDate_(dateString) {

  if (
    !dateString
  ) {

    throw new Error(
      "Tanggal belum diisi."
    );

  }

  const parts =
    String(dateString).split("-");

  if (
    parts.length !== 3
  ) {

    throw new Error(
      "Format tanggal tidak valid: " +
      dateString
    );

  }

  const year =
    Number(parts[0]);

  const month =
    Number(parts[1]) - 1;

  const day =
    Number(parts[2]);

  return new Date(
    year,
    month,
    day
  );

}

/************************************************
 * PARSE DATE TIME
 * Input date: yyyy-mm-dd
 * Input time: HH:mm
 ************************************************/

function parseDateTime_(
  dateString,
  timeString
) {

  if (
    !dateString
  ) {

    throw new Error(
      "Tanggal belum diisi."
    );

  }

  const dateParts =
    String(dateString).split("-");

  if (
    dateParts.length !== 3
  ) {

    throw new Error(
      "Format tanggal tidak valid: " +
      dateString
    );

  }

  const time =
    timeString && String(timeString).trim() !== ""
      ? String(timeString)
      : "00:00";

  const timeParts =
    time.split(":");

  const year =
    Number(dateParts[0]);

  const month =
    Number(dateParts[1]) - 1;

  const day =
    Number(dateParts[2]);

  const hour =
    Number(timeParts[0]) || 0;

  const minute =
    Number(timeParts[1]) || 0;

  return new Date(
    year,
    month,
    day,
    hour,
    minute,
    0
  );

}

/************************************************
 * TAMBAH HARI
 ************************************************/

function addDays_(
  date,
  days
) {

  const newDate =
    new Date(date);

  newDate.setDate(
    newDate.getDate() + days
  );

  return newDate;

}

/************************************************
 * HITUNG SELISIH HARI
 ************************************************/

function daysDifference_(
  dateA,
  dateB
) {

  const oneDay =
    24 * 60 * 60 * 1000;

  const a =
    new Date(
      dateA.getFullYear(),
      dateA.getMonth(),
      dateA.getDate()
    );

  const b =
    new Date(
      dateB.getFullYear(),
      dateB.getMonth(),
      dateB.getDate()
    );

  return Math.round(
    (b - a) / oneDay
  );

}

/************************************************
 * AMBIL / BUAT SHEET
 ************************************************/

function getLogSheet_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (
    !ss
  ) {

    throw new Error(
      "Spreadsheet tidak ditemukan. Pastikan script ini terhubung dengan Google Sheet."
    );

  }

  let sheet =
    ss.getSheetByName(
      SHEET_NAME
    );

  if (
    !sheet
  ) {

    sheet =
      ss.insertSheet(
        SHEET_NAME
      );

  }

  if (
    sheet.getLastRow() === 0
  ) {

    sheet.appendRow([

      "No",
      "Title",
      "Start Date",
      "Start Time",
      "End Time",
      "End Date",
      "All Day",
      "Repeat",
      "Repeat Interval",
      "Repeat Unit",
      "Repeat Days",
      "Repeat End Type",
      "Repeat End Date",
      "Repeat Count",
      "Location",
      "Reminder",
      "Guests",
      "Description",
      "Send Notification",
      "Status",
      "Event ID",
      "Created At"

    ]);

  }

  return sheet;

}

/************************************************
 * WEEKDAY MAP
 ************************************************/

function getWeekdayMap_() {

  return {

    MON:
      CalendarApp.Weekday.MONDAY,

    TUE:
      CalendarApp.Weekday.TUESDAY,

    WED:
      CalendarApp.Weekday.WEDNESDAY,

    THU:
      CalendarApp.Weekday.THURSDAY,

    FRI:
      CalendarApp.Weekday.FRIDAY,

    SAT:
      CalendarApp.Weekday.SATURDAY,

    SUN:
      CalendarApp.Weekday.SUNDAY

  };

}

/************************************************
 * GET WEEKDAY CODE FROM DATE
 ************************************************/

function getWeekdayCodeFromDate_(date) {

  const codes = [

    "SUN",
    "MON",
    "TUE",
    "WED",
    "THU",
    "FRI",
    "SAT"

  ];

  return codes[
    date.getDay()
  ];

}

/************************************************
 * GET REPEAT DAY CODES
 ************************************************/

function getRepeatDayCodes_(form) {

  if (
    !form.repeatDays ||
    String(form.repeatDays).trim() === ""
  ) {

    return [];

  }

  return String(form.repeatDays)
    .split(",")
    .map(day => day.trim().toUpperCase())
    .filter(Boolean);

}

/************************************************
 * CARI TANGGAL PERTAMA YANG SESUAI WEEKLY
 ************************************************/

function getFirstMatchingWeeklyDate_(
  startDate,
  repeatDayCodes
) {

  if (
    !repeatDayCodes ||
    repeatDayCodes.length === 0
  ) {

    return startDate;

  }

  const checkDate =
    new Date(startDate);

  for (
    let i = 0;
    i < 7;
    i++
  ) {

    const dayCode =
      getWeekdayCodeFromDate_(
        checkDate
      );

    if (
      repeatDayCodes.indexOf(dayCode) !== -1
    ) {

      return checkDate;

    }

    checkDate.setDate(
      checkDate.getDate() + 1
    );

  }

  return startDate;

}

/************************************************
 * ADJUST WEEKLY DATE
 * Supaya tidak muncul event H lalu repeat H+1
 ************************************************/

function getAdjustedWeeklyDates_(
  form,
  start,
  end,
  startDateOnly
) {

  const result = {

    start:
      start,

    end:
      end,

    startDateOnly:
      startDateOnly

  };

  if (
    form.repeat !== "Y"
  ) {

    return result;

  }

  if (
    String(form.repeatUnit || "").toUpperCase() !== "WEEKLY"
  ) {

    return result;

  }

  const repeatDayCodes =
    getRepeatDayCodes_(
      form
    );

  if (
    repeatDayCodes.length === 0
  ) {

    return result;

  }

  const firstWeeklyDate =
    getFirstMatchingWeeklyDate_(
      startDateOnly,
      repeatDayCodes
    );

  const dayShift =
    daysDifference_(
      startDateOnly,
      firstWeeklyDate
    );

  result.startDateOnly =
    firstWeeklyDate;

  result.start =
    addDays_(
      start,
      dayShift
    );

  result.end =
    addDays_(
      end,
      dayShift
    );

  return result;

}

/************************************************
 * BUILD RECURRENCE
 ************************************************/

function buildRecurrence_(form) {

  const recurrence =
    CalendarApp.newRecurrence();

  const repeatUnit =
    String(
      form.repeatUnit || ""
    ).toUpperCase();

  const interval =
    Math.max(
      1,
      Number(form.repeatInterval) || 1
    );

  let rule = null;

  /**********************************************
   * REPEAT UNIT
   **********************************************/

  switch (
    repeatUnit
  ) {

    case "DAILY":

      rule =
        recurrence.addDailyRule();

      break;

    case "WEEKLY":

      rule =
        recurrence.addWeeklyRule();

      if (
        form.repeatDays &&
        String(form.repeatDays).trim() !== ""
      ) {

        const weekdayMap =
          getWeekdayMap_();

        const selectedWeekdays =
          String(form.repeatDays)
            .split(",")
            .map(day => day.trim().toUpperCase())
            .map(day => weekdayMap[day])
            .filter(Boolean);

        if (
          selectedWeekdays.length === 1
        ) {

          rule.onlyOnWeekday(
            selectedWeekdays[0]
          );

        }

        else if (
          selectedWeekdays.length > 1
        ) {

          rule.onlyOnWeekdays(
            selectedWeekdays
          );

        }

      }

      break;

    case "MONTHLY":

      rule =
        recurrence.addMonthlyRule();

      break;

    case "YEARLY":

      rule =
        recurrence.addYearlyRule();

      break;

    default:

      throw new Error(
        "Repeat unit tidak valid. Nilai yang diterima: " +
        form.repeatUnit
      );

  }

  /**********************************************
   * REPEAT INTERVAL
   **********************************************/

  rule.interval(
    interval
  );

  /**********************************************
   * REPEAT END TYPE
   **********************************************/

  if (
    form.repeatEndType === "DATE" &&
    form.repeatEndDate
  ) {

    const untilDate =
      parseDateTime_(
        form.repeatEndDate,
        "23:59"
      );

    rule.until(
      untilDate
    );

  }

  if (
    form.repeatEndType === "COUNT" &&
    form.repeatCount
  ) {

    const count =
      Number(form.repeatCount);

    if (
      isNaN(count) ||
      count < 1
    ) {

      throw new Error(
        "Jumlah occurrence repeat tidak valid."
      );

    }

    rule.times(
      count
    );

  }

  return recurrence;

}

/************************************************
 * BUILD OPTIONS EVENT
 ************************************************/

function buildEventOptions_(form) {

  validateEmails(
    form.guests
  );

  const guests =
    normalizeGuests_(
      form.guests
    );

  const options = {

    description:
      form.description || "",

    location:
      form.location || ""

  };

  /**********************************************
   * GUEST
   * Guest tetap masuk meskipun send notification = N
   **********************************************/

  if (
    guests
  ) {

    options.guests =
      guests;

    options.sendInvites =
      form.sendNotification === "Y";

  }

  return options;

}

/************************************************
 * VALIDASI FORM DASAR
 ************************************************/

function validateForm_(form) {

  if (
    !form
  ) {

    throw new Error(
      "Data form tidak ditemukan."
    );

  }

  if (
    !form.title ||
    String(form.title).trim() === ""
  ) {

    throw new Error(
      "Judul event wajib diisi."
    );

  }

  if (
    !form.startDate
  ) {

    throw new Error(
      "Start Date wajib diisi."
    );

  }

  if (
    !form.endDate
  ) {

    throw new Error(
      "End Date wajib diisi."
    );

  }

  if (
    form.repeat === "Y"
  ) {

    if (
      !form.repeatUnit
    ) {

      throw new Error(
        "Repeat unit wajib dipilih."
      );

    }

    if (
      String(form.repeatUnit).toUpperCase() === "WEEKLY" &&
      (
        !form.repeatDays ||
        String(form.repeatDays).trim() === ""
      )
    ) {

      throw new Error(
        "Untuk repeat weekly, pilih minimal satu hari."
      );

    }

  }

}

/************************************************
 * SIMPAN EVENT
 ************************************************/

function saveEvent(form) {

  try {

    Logger.log(
      JSON.stringify(form)
    );

    validateForm_(
      form
    );

    const sheet =
      getLogSheet_();

    const calendar =
      CalendarApp
        .getDefaultCalendar();

    const options =
      buildEventOptions_(
        form
      );

    const isAllDay =
      form.allDay === "Y";

    const isRepeat =
      form.repeat === "Y";

    const start =
      parseDateTime_(
        form.startDate,
        form.startTime || "00:00"
      );

    const end =
      parseDateTime_(
        form.endDate,
        form.endTime || "23:59"
      );

    const startDateOnly =
      parseDate_(
        form.startDate
      );

    const endDateOnly =
      parseDate_(
        form.endDate
      );

    const adjustedWeekly =
      getAdjustedWeeklyDates_(
        form,
        start,
        end,
        startDateOnly
      );

    const adjustedStart =
      adjustedWeekly.start;

    const adjustedEnd =
      adjustedWeekly.end;

    const adjustedStartDateOnly =
      adjustedWeekly.startDateOnly;

    let event = null;

    /**********************************************
     * VALIDASI WAKTU UNTUK EVENT BIASA
     **********************************************/

    if (
      !isAllDay &&
      end <= start
    ) {

      throw new Error(
        "End Date/Time harus lebih besar dari Start Date/Time."
      );

    }

    /**********************************************
     * FULL DAY + REPEAT
     **********************************************/

    if (
      isAllDay &&
      isRepeat
    ) {

      const recurrence =
        buildRecurrence_(
          form
        );

      event =
        calendar.createAllDayEventSeries(

          form.title,

          adjustedStartDateOnly,

          recurrence,

          options

        );

    }

    /**********************************************
     * FULL DAY TANPA REPEAT
     **********************************************/

    else if (
      isAllDay
    ) {

      if (
        endDateOnly > startDateOnly
      ) {

        const exclusiveEndDate =
          addDays_(
            endDateOnly,
            1
          );

        event =
          calendar.createAllDayEvent(

            form.title,

            startDateOnly,

            exclusiveEndDate,

            options

          );

      }

      else {

        event =
          calendar.createAllDayEvent(

            form.title,

            startDateOnly,

            options

          );

      }

    }

    /**********************************************
     * EVENT BIASA + REPEAT
     **********************************************/

    else if (
      isRepeat
    ) {

      const recurrence =
        buildRecurrence_(
          form
        );

      event =
        calendar.createEventSeries(

          form.title,

          adjustedStart,

          adjustedEnd,

          recurrence,

          options

        );

    }

    /**********************************************
     * EVENT BIASA TANPA REPEAT
     **********************************************/

    else {

      event =
        calendar.createEvent(

          form.title,

          start,

          end,

          options

        );

    }

    /**********************************************
     * REMINDER
     **********************************************/

    if (
      form.notification &&
      !isNaN(form.notification)
    ) {

      const reminderMinutes =
        Number(
          form.notification
        );

      if (
        reminderMinutes >= 0
      ) {

        try {

          event.addPopupReminder(
            reminderMinutes
          );

        } catch (err) {

          Logger.log(
            "Reminder error: " + err
          );

        }

      }

    }

    /**********************************************
     * SIMPAN KE SHEET
     **********************************************/

    const nextNo =
      Math.max(
        1,
        sheet.getLastRow()
      );

    const eventId =
      event && event.getId
        ? event.getId()
        : "";

    sheet.appendRow([

      nextNo,

      form.title || "",

      form.startDate || "",

      form.startTime || "",

      form.endTime || "",

      form.endDate || "",

      form.allDay || "N",

      form.repeat || "N",

      form.repeatInterval || "",

      form.repeatUnit || "",

      form.repeatDays || "",

      form.repeatEndType || "",

      form.repeatEndDate || "",

      form.repeatCount || "",

      form.location || "",

      form.notification || "",

      form.guests || "",

      form.description || "",

      form.sendNotification || "N",

      "CREATED",

      eventId,

      new Date()

    ]);

    return {

      success:
        true,

      message:
        "Event berhasil dibuat",

      eventId:
        eventId

    };

  } catch (err) {

    Logger.log(
      err
    );

    throw new Error(
      err.message
    );

  }

}
