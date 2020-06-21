#include <Oxygen.h>
#include <MutichannelGasSensor.h>
#include <math.h>
#include <SoftwareSerial.h>
#include "rgb_lcd.h"
#include <SPI.h>

#define MAX6675_CS   10
#define MAX6675_SO   12
#define MAX6675_SCK  13

rgb_lcd lcd;

const int colorR = 255;
const int colorG = 255;
const int colorB = 0;
// Grove - Gas Sensor(O2) test code

const float VRefer = 5;       // voltage of adc reference

const int pinAdc   = A0;

void setup() 
{
  lcd.begin(16, 2);

    lcd.setRGB(colorR, colorG, colorB);

    Serial.begin(9600);
   //Serial.println("This Program just for oxygen test");
   
}

void loop() 
{
  float temperature_read = readThermocouple(); 
    float Vout =0;
  lcd.setCursor(0, 0);
     
  //lcd.setCursor(0,0);
  lcd.print("T : ");
 // lcd.setCursor(7,1);  
  lcd.print(temperature_read-5,1); 
  lcd.print((char)223);
  lcd.print("C");
  Serial.print("Temp: ");
    Serial.print(readThermocouple());
    Serial.println((char) 223);
    Vout = readO2Vout();
    //Serial.print("Vout:");
    //Serial.print(Vout);
    //Serial.print("V");
    Serial.print("O2_Percentage: ");
    Serial.println(readConcentration());
    lcd.setCursor(0, 1);
    lcd.print("O2: ");
    lcd.print(readConcentration());
    lcd.print("%");
   // Serial.println("%");
    delay(1000);
}

double readThermocouple() {

  uint16_t v;
  pinMode(MAX6675_CS, OUTPUT);
  pinMode(MAX6675_SO, INPUT);
  pinMode(MAX6675_SCK, OUTPUT);
  
  digitalWrite(MAX6675_CS, LOW);
  delay(1);

  // Read in 16 bits,
  //  15    = 0 always
  //  14..2 = 0.25 degree counts MSB First
  //  2     = 1 if thermocouple is open circuit  
  //  1..0  = uninteresting status
  
  v = shiftIn(MAX6675_SO, MAX6675_SCK, MSBFIRST);
  v <<= 8;
  v |= shiftIn(MAX6675_SO, MAX6675_SCK, MSBFIRST);
  
  digitalWrite(MAX6675_CS, HIGH);
  if (v & 0x4) 
  {    
    // Bit 2 indicates if the thermocouple is disconnected
    return NAN;     
  }

  // The lower three bits (0,1,2) are discarded status bits
  v >>= 3;

  // The remaining bits are the number of 0.25 degree (C) counts
  return v*0.25;
}

float readO2Vout()
{
    long sum = 0;
    for(int i=0; i<50; i++)
    {
        sum += analogRead(pinAdc);
    }
    
    sum >>= 5;
    
    float MeasuredVout = sum * (VRefer / 1023.0);
    return MeasuredVout;
}

float readConcentration()
{
    float MeasuredVout = readO2Vout();
    float Concentration = MeasuredVout * 0.21 / 2.0;
    float Concentration_Percentage=Concentration*100;
    return Concentration_Percentage;
}
