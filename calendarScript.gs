function createWeeklyEventsFromTemplate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  const baseWeek = new Date('2024-10-06'); // Start from Sunday, October 6, 2024

  const dayMap = {
    "SUNDAY": 0,
    "MONDAY": 1,
    "TUESDAY": 2,
    "WEDNESDAY": 3,
    "THURSDAY": 4,
    "FRIDAY": 5,
    "SATURDAY": 6,
  };

  const colorMap = {
    "Family Time": "9",
    "Personal Growth": "10",
    "Family Lunch": "9",
    "Relaxation Time": "5",
    "Business Planning": "7",
    "Buffer Time": "2",
    "Family Dinner": "9",
    "Personal Care": "2",
    "Breakfast": "3",
    "Patient Appointments": "4",
    "Break": "5",
    "Lunch Break": "5",
    "Quick Team Check-in or Meeting Block": "8",
    "Meeting Block": "8",
    "Content Creation": "6",
    "Daily Wrap-Up": "10",
    "Marketing Strategy": "6",
    "Family Activity": "9",
    "Planning": "7",
    "Marketing Content": "6",
    "Strategic Planning": "7",
    "Email Management": "4",
    "Weekly Wrap-Up": "7",
    "Family Breakfast": "9",
    "Family Outing": "9",
    "Resting Time": "5",
    "Planning Session": "7"
  };

  let currentDay = "";

  for (let i = 3; i < data.length; i++) {
    let [ , day, timeRange, title, , description] = data[i];

    if (day && day.trim() !== "") {
      currentDay = day.trim();
    }

    if (!timeRange || !title || !currentDay) continue;

    const dayIndex = dayMap[currentDay.toUpperCase()];
    if (dayIndex === undefined) continue;

    const [startTimeStr, endTimeStr] = timeRange.split(" - ").map(str => str.trim());

    const startDateTime = new Date(baseWeek);
    startDateTime.setDate(baseWeek.getDate() + dayIndex);
    startDateTime.setHours(...convertTo24Hour(startTimeStr));

    const endDateTime = new Date(baseWeek);
    endDateTime.setDate(baseWeek.getDate() + dayIndex);
    endDateTime.setHours(...convertTo24Hour(endTimeStr));

    const event = calendar.createEvent(title, startDateTime, endDateTime, {
      description: description || "",
      recurrence: CalendarApp.newRecurrence().addWeeklyRule().times(10)
    });

    if (colorMap[title]) {
      event.setColor(colorMap[title]);
    }
  }
}

function convertTo24Hour(timeStr) {
  const [time, modifier] = timeStr.toLowerCase().split(" ");
  let [hours, minutes] = time.split(":".replace(/\u200B/g, '')).map(Number);
  if (modifier === "pm" && hours !== 12) hours += 12;
  if (modifier === "am" && hours === 12) hours = 0;
  return [hours, minutes];
}
