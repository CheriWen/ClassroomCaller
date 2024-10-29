#include <LedControl.h>

const int DIN_PIN = 5;
const int CLK_PIN = 3;
const int CS_PIN = 4;
const int NUM_DEVICES = 4; // 8x32 点阵屏有4块8x8的矩阵

LedControl lc = LedControl(DIN_PIN, CLK_PIN, CS_PIN, NUM_DEVICES);
String message = "";  // 保存要显示的消息
int scrollSpeed = 100; // 滚动速度，值越小速度越快

void setup() {
    Serial.begin(9600); 
    for (int i = 0; i < NUM_DEVICES; i++) {
        lc.shutdown(i, false);  // 打开显示
        lc.setIntensity(i, 8);  // 设置亮度（0~15）
        lc.clearDisplay(i);     // 清空显示
    }
    Serial.println("Please enter a message to display:");
}

void loop() {
    // 接收串口输入内容
    while (Serial.available() > 0) {
        char received = Serial.read();
        if (received == '\n') {
            displayMessage(message);
            message = "";
        } else {
            message += received;
        }
    }
}

void displayMessage(String msg) {
    // 滚动显示消息
    int len = msg.length();
    for (int pos = 0; pos < len * 8 + 32; pos++) {
        lc.clearDisplay(0);
        for (int i = 0; i < len; i++) {
            int charColumn = pos - i * 8;
            if (charColumn >= -8 && charColumn < 32) {
                displayChar(msg[i], charColumn);
            }
        }
        delay(scrollSpeed);
    }
}

// 显示单个字符
void displayChar(char c, int column) {
    byte charData[8];
    getCharData(c, charData);

    for (int col = 0; col < 8; col++) {
        int displayColumn = column + col;
        if (displayColumn >= 0 && displayColumn < 32) {
            int device = displayColumn / 8;
            int colInDevice = displayColumn % 8;
            lc.setColumn(device, colInDevice, charData[col]);
        }
    }
}

// 获取字符的点阵数据
void getCharData(char c, byte data[8]) {
    // 定义字符点阵（仅作为示例，可根据需求自行扩展）
    byte font[][8] = {
        {0x00, 0x3E, 0x51, 0x49, 0x45, 0x3E, 0x00, 0x00},  // '0'
        {0x00, 0x00, 0x42, 0x7F, 0x40, 0x00, 0x00, 0x00},  // '1'
        {0x00, 0x42, 0x61, 0x51, 0x49, 0x46, 0x00, 0x00},  // '2'
        // 可继续添加其他字符的数据
    };

    if (c >= '0' && c <= '2') {
        for (int i = 0; i < 8; i++) {
            data[i] = font[c - '0'][i];
        }
    } else {
        for (int i = 0; i < 8; i++) {
            data[i] = 0x00;
        }
    }
}
