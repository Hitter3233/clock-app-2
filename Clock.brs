' ============================================================
' ROKU CLOCK SETTINGS
' ============================================================
'
' Roku screen resolution:
' 1920 x 1080
'
' Change these numbers to move the clock and date.
'
' X:
'   Higher number = move RIGHT
'   Lower number  = move LEFT
'
' Y:
'   Higher number = move DOWN
'   Lower number  = move UP
'
' ============================================================


' ------------------------------------------------------------
' TIME POSITION
' ------------------------------------------------------------

TIME_X = 640
TIME_Y = 400


' ------------------------------------------------------------
' DATE POSITION
' ------------------------------------------------------------

DATE_X = 640
DATE_Y = 500


' ============================================================
' CLOCK SETUP
' ============================================================

sub SetupClock(clockLabel, dateLabel)

    ' Move the time
    clockLabel.translation = [TIME_X, TIME_Y]

    ' Move the date
    dateLabel.translation = [DATE_X, DATE_Y]

end sub


' ============================================================
' UPDATE CLOCK DISPLAY
' ============================================================

sub UpdateClockDisplay(clockLabel, dateLabel)

    dt = CreateObject("roDateTime")
    dt.ToLocalTime()


    ' --------------------------------------------------------
    ' TIME
    ' --------------------------------------------------------

    hours = dt.GetHours()
    minutes = dt.GetMinutes()
    seconds = dt.GetSeconds()

    ' Determine AM / PM
    if hours >= 12
        ampm = "PM"
    else
        ampm = "AM"
    end if

    ' Convert 24-hour time to 12-hour time
    if hours = 0
        displayHour = 12
    else if hours > 12
        displayHour = hours - 12
    else
        displayHour = hours
    end if

    ' Add leading zero to minutes
    if minutes < 10
        minuteText = "0" + minutes.ToStr()
    else
        minuteText = minutes.ToStr()
    end if

    ' Add leading zero to seconds
    if seconds < 10
        secondText = "0" + seconds.ToStr()
    else
        secondText = seconds.ToStr()
    end if

    ' Create time string
    clockLabel.text = displayHour.ToStr() + ":" + minuteText + ":" + secondText + " " + ampm


    ' --------------------------------------------------------
    ' DATE
    ' --------------------------------------------------------

    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December"
    ]

    monthNumber = dt.GetMonth()

    monthName = months[monthNumber - 1]

    day = dt.GetDayOfMonth()
    year = dt.GetYear()

    ' Create date string
    dateLabel.text = monthName + " " + day.ToStr() + ", " + year.ToStr()

end sub