#include <stdlib.h>
#include <stdio.h>

int main(){
unsigned char sc[] = %s;
unsigned char key = %s;
unsigned int SC_LEN = %s;
int i;
unsigned char* tmp = (unsigned char*)malloc(SC_LEN);
%s
for(i=0; i<SC_LEN; i++){
    %s
    %s
    %s
}
%s
((void (*)())tmp)();
%s
return 0;
}
