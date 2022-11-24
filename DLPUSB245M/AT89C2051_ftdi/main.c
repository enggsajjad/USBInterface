#include<reg2051.h>
#include<intrins.h> 
sbit RXF = P3^2;
sbit TXE = P3^3;
sbit RD = P3^4;
sbit WR = P3^5;

bit rxflag;
bit flag;
bit Txflag;
unsigned char rxChr;
unsigned char cnt;
void Delay25ms(void)
{
	TL0 = 0x4C;//25ms
	TH0 = 0x00;
	TR0 = 1;
	while(!TF0);
	TR0 = 0;
	TF0 = 0;
}
void PrintChar(unsigned char c)
{

	TI=1;
	while (!TI);	TI=0;	SBUF = c;
	while (!TI);	TI=0;
}
void PrintString(char *s)
{
	while (*s)
	{
		PrintChar(*s);
		s++;
	}
}
unsigned char DLP_Read(void)
{
	unsigned char q;
	if(!RXF)
	{
		P1 = 0xFF;
		//_nop_();
		//_nop_();
		RD = 0; 
		q = P1;
		//_nop_();
		//_nop_(); 
		RD = 1;
		return q;
	}
}

void DLP_Send(unsigned char s)
{
	if(!TXE)
	{
	 	 P1 = 0x00;
	 	 //_nop_(); 
	 	 P1 = s;	
	 	 _nop_();
	 	 WR = 0; 
	 	 //_nop_();
		 //_nop_();
	 	 WR = 1;
	 	 P1 = 0xFF; 
	}
}

void main(void)
{
	unsigned char temp;
	unsigned char cnt=0;
	SCON = 0x50;
	TMOD = 0x21;
	TL1 = 0xFD;
	TH1 = 0xFD;
	TR1 = 1;
	EX0 = 1;
	IT0 = 1;
	ES	= 1;
	EA = 1;
	RD = 1;
	WR = 1;
	
	while(1)
	{
	 	if(flag)
	 	{
	 		flag = 0;
	 		temp = DLP_Read();
	 		//PrintChar(temp);
	 		Txflag = 1;
	 	}
	 	if(Txflag)
	 	{
	 		Txflag = 0;
	 		DLP_Send(temp);
	 	}
	}//while
}//main
void Recieve_EX0() interrupt 0
{
 	flag = 1;
}
void Serial() interrupt 4
{
	if (RI)
	{
		RI = 0;
  		rxChr = SBUF;
  		rxflag = 1;
  		switch(rxChr)
  		{
  			case 'A':
  				DLP_Send('A');
  				DLP_Send('B');
  				DLP_Send('C');
  				DLP_Send('D');
  				DLP_Send('E');
  				DLP_Send('F');
  				break;
  			case 'B':
  				for(cnt=0;cnt<225;cnt++)
  					DLP_Send(cnt);
  				for(cnt=0;cnt<225;cnt++)
  					DLP_Send(cnt);
  				for(cnt=0;cnt<225;cnt++)
  					DLP_Send(cnt);
  				for(cnt=0;cnt<225;cnt++)
  					DLP_Send(cnt);
  				break;
  			default:
  				break;
  		}
  	}//End if
}




































