// 2023-09-13
// pista automobiline, alimentazione e comunicazione 
// tramite cavo UDB
// 2 ingressi digitali (uno per ogni macchinina)
// viene trasmesso il passaggio della/e macchinina/e con i codici 1,2,3
// filtrando l'ingresso per 1 secondo (filtro_ms)

#include "Arduino_LED_Matrix.h"

// configurazione
const int macchina_A = 8;
const int macchina_B = 7;
const int filtro_ms = 1000;

// ram di lavoro per display
uint32_t LED_RAM[] = {
	0x00000000,
	0x00000000,
	0x00000000
};

// pattern per display
uint32_t LED_WELCOME[] = {
  0xfa18b38a,
  0xd8a1fa18,
  0x21821821
};
uint32_t LED_VERSION[] = {
  0x31ff1,
  0xb31b31b3,
  0x1b35f000
};
const uint32_t LED_CLEAN[] = {
	0x00000000,
	0x00000000,
	0x00000000
};
const uint32_t LED_A[] = {
  0x20050088,
  0x880f808,
  0x80880880
};
const uint32_t LED_AB[] = {
  0x20e50988,
  0x988ef8e8,
  0x8988988e
};
const uint32_t LED_B[] = {
  0xe00900,
  0x900e00e0,
  0x900900e
};

uint32_t LED_OK[] = {
  0x7a44a,
  0x84b04b04,
  0xa87a4000
};

// variabili di lavoro
unsigned long tempo_a;
unsigned long tempo_b;
int valore_a;
int valore_b;
int run;

// istanza classe di libreria
ArduinoLEDMatrix matrix;

// inizializzazione
void setup() {
  Serial.begin(115200);
  pinMode(macchina_A, INPUT);
  pinMode(macchina_B, INPUT);
  matrix.begin();
  matrix.loadFrame(LED_WELCOME);
  delay(2000);
  matrix.loadFrame(LED_VERSION);
  delay(2000);
  tempo_a  = 0;
  tempo_b  = 0;
  valore_a = 0; 
  valore_b = 0; 
  run      = 0;

  while (!Serial) {
    ; // attesa seriale (USB) pronta
  }
  matrix.loadFrame(LED_OK);
}

// chiamata a loop
void loop() {
  int status = 0;
  unsigned long tempo_adesso = millis();
  int a_adesso = digitalRead(macchina_A);
  int b_adesso = digitalRead(macchina_B);
  
  // prolungamento valori per filtraggio rimbalzi
  if (a_adesso) { 
    tempo_a = tempo_adesso; 
    if(valore_a == 0) {
      status  |= 1;
    }
    valore_a = 1; 
  } else if (tempo_adesso > tempo_a + filtro_ms) {
    valore_a = 0; 
  }
  if (b_adesso) { 
      tempo_b = tempo_adesso; 
      if(valore_b == 0) {
        status  |= 2;
      }
      valore_b = 1; 
  } else if (tempo_adesso > tempo_b + filtro_ms) {
    valore_b = 0; 
  }
  // trasmissione sul fronte di salita
  if(status == 1) {
    Serial.write('1');
  } else if (status == 2) {
    Serial.write('2');
  } else if (status == 3) {
    Serial.write('3');
  }
  // gestione display
  if (valore_a && !valore_b) {
    matrix.loadFrame(LED_A);
    run = 1;
  } else if (!valore_a && valore_b) {
    matrix.loadFrame(LED_B);
    run = 1;
  } else if (valore_a && valore_b) {
    run = 1;
    matrix.loadFrame(LED_AB);
  } else if (run == 1) {
    matrix.loadFrame(LED_CLEAN);
  }
}
