import { CalendarDate, getLocalTimeZone } from "@internationalized/date";

export class Util {
  static formattedDate(dateCreated?: Date | string): string {
    const _dateCreated = dateCreated || new Date();

    const date = new Date(_dateCreated);

    const userLocale = navigator.language || "en-US";
    const formattedDate = date.toLocaleString(userLocale, {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });

    return formattedDate;
  }

  static dateToCalendarDate(date?: Date): CalendarDate {
    const _date = date || new Date();

    return new CalendarDate(
      _date.getFullYear(),
      _date.getMonth() + 1,
      _date.getDate(),
    );
  }

  static calendarToDate(calendarDate?: CalendarDate | null | undefined): Date {
    const now = new Date();
    const _calendarDate =
      calendarDate ||
      new CalendarDate(now.getFullYear(), now.getMonth() + 1, now.getDate());

    return _calendarDate.toDate(getLocalTimeZone() || "en-US");
  }
}
