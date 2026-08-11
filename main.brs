sub Main()
    screen = CreateObject("roSGScreen")
    scene = screen.CreateScene("MainScene")

    port = CreateObject("roMessagePort")
    screen.SetMessagePort(port)

    screen.Show()

    while true
        msg = wait(0, port)

        if type(msg) = "roSGScreenEvent"
            if msg.IsScreenClosed()
                exit while
            end if
        end if
    end while
end sub