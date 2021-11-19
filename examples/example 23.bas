'EXAMPLE #23 - Inkey Variable Test

window "win", "Inkey Test", normal, 100, 100, 500, 400

event "win", "keyup", KeyPress

pause


sub KeyPress
    message Inkey, ""
end sub