sub init()

    ' Get the clock and date objects
    m.clockLabel = m.top.findNode("clockLabel")
    m.dateLabel = m.top.findNode("dateLabel")

    ' Get the background
    m.background = m.top.findNode("background")

    ' Set up the clock positions
    SetupClock(m.clockLabel, m.dateLabel)

    ' Start the clock timer
    m.timer = CreateObject("roSGNode", "Timer")
    m.timer.duration = 1
    m.timer.repeat = true
    m.timer.observeField("fire", "UpdateClock")
    m.top.appendChild(m.timer)

    ' Update immediately
    UpdateClock()

    ' Start timer
    m.timer.control = "start"

end sub


sub UpdateClock()

    UpdateClockDisplay(m.clockLabel, m.dateLabel)

end sub