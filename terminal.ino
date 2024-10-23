#include <LiquidCrystal.h>

// Initialize the library with the numbers of the interface pins
LiquidCrystal lcd(12, 11, 5, 4, 3, 2);

void setup() {
    Serial.begin(9600);
    lcd.begin(16, 2);
    lcd.print("Happy Study!");
}

void loop() {
    if (Serial.available()){
        String data = Serial.readStringUntil('\n');
        Serial.flush();

        //Clear the LCD
        lcd.clear();
        lcd.setCursor(0, 0);
        lcd.print("Please call:");
        lcd.setCursor(0,1);
        lcd.print(data);
    }
}