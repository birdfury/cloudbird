import os
import xlwings as xw
import argparse

def data2_excel(a,b,c):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    sht = wb.sheets[0]
    sht.name = 'first'

    sht.range('A1').value = '内存地址'
    sht.range('B1').value = '目的地址'
    sht.range('C1').value = '地址偏移'

    for i in range(len(a)):
        sht.range('A' + str(i+2)).value = a[i]
        sht.range('B' + str(i+2)).value = b[i]
        sht.range('C' + str(i+2)).value = c[i]
    wb.save('CalculateJumpAddress00Copy.xlsx')
    wb.close()
    app.quit()


class calculate_armJumpInstruction_address(object):
    def __init__(self, filename, startaddr):
        self.localladdr = []
        self.dstaddr = []
        self.offsetaddr = []
        self.filename = filename
        self.startaddr = startaddr

    def judge_jump_instruction(self):
        binfile = open(self.filename, 'rb')
        binfile.seek(int(self.startaddr), 0)

        Size = os.path.getsize(self.filename)
        size = Size - self.startaddr
        for i in range(0, int(size/4)):
            localaddress = i * 4 + self.startaddr

            flags = 0
            buffer = binfile.read(4)
            last = int.from_bytes(buffer, 'little')
            #print('{:02x}'.format(last))
            bit1 = last & 0xff
            bit2 = (last & 0xff00) >> 8
            bit3 = (last & 0xff0000) >> 16
            bit4 = (last & 0xff000000) >> 24
            if ( (bit4 != b'') and (bit4 & 0xff == 0xeb) and ((bit2 & 0xf8) != 0xf0) and ((bit3 & 0xf8) != 0xf0)):
                print('ARM BL local address is ' + hex(localaddress))
                flags = 1                       #ARMBL
                self.calculate_armBL_addr(last, localaddress)
            if ( (bit4 != b'') and ((bit4 & 0xfe) == 0xfa) and ((bit2 & 0xf8) != 0xf0) and ((bit3 & 0xf8) != 0xf0)):
                print('ARM BLX local address is ' + hex(localaddress))
                flags = 2                       #ARMBLX
                self.calculate_armBLX_addr(last, localaddress)
            else:
                flags = 0
                continue

        binfile.close()
    def write_excel(self):
        data2_excel(self.localladdr, self.dstaddr, self.offsetaddr)


    def calculate_armBL_addr(self, binfile, localAddress):
        machineNumber = (binfile & 0xff0000 << 16) + (binfile & 0xff00 << 8) + (binfile & 0xff)
        if (((machineNumber & 0xffffff) >> 23) == 1):
            machineNumber |= 0xff0000
        if (((machineNumber & 0xffffff) >> 23) == 0):
            machineNumber &= 0x00ffffff
        machineNumberTo32 = (machineNumber << 2) & 0xffffffff
        machineNumberAddFlag = machineNumberTo32 + (((binfile & 0xff000000) & 0x1) << 1)

        destinationAddress = localAddress + machineNumberAddFlag + 0x8
        destinationAddress &= 0xffffffff
        print('This is a armBl command')
        print('ARM BL destination address is ' + hex(destinationAddress))

        offsetAddress = abs(destinationAddress - localAddress)

        print('ARM BL Address offset is ' + hex(offsetAddress))
        self.localladdr.append(localAddress)
        self.dstaddr.append(destinationAddress)
        self.offsetaddr.append(offsetAddress)



    def calculate_armBLX_addr(self, HEXfile, localAddr):

        machineCode = (HEXfile & 0xff0000 << 16) + (HEXfile & 0xff00 << 8) + (HEXfile & 0xff)

        if ((machineCode & 0xffffff) >> 23 == 1):
            machineCode |= 0xff000000
        if ((machineCode & 0xffffff) >> 23 == 0):
            machineCode &= 0x00efffff

        machineCodeTo32 = ((machineCode << 2) & 0xffffffff)
        machineCodeAdd = (machineCodeTo32 + ((HEXfile & 0xff000000) & 0x1) << 1)

        destAddr = localAddr + machineCodeAdd + 0x8
        destAddr &= 0xffffffff
        print('This is a armBLX command')
        print('ARM BLX destination address is ' + hex(destAddr))

        offsetAddr = abs(destAddr - localAddr)

        print('ARM BLX offset address is ' + hex(offsetAddr))

        self.localladdr.append(localAddr)
        self.dstaddr.append(destAddr)
        self.offsetaddr.append(offsetAddr)



class calculate_thumbInstruction_jumpAddr():

    def __init__(self,BINFile,offsetAddr):
        self.localladdr = []
        self.dstaddr = []
        self.offsetaddr = []
        self.BINFile = BINFile
        self.offsetAddr = offsetAddr
    def judge_jump_instruction(self):
        binfile = open(self.BINFile, 'rb')
        binfile.seek(int(self.offsetAddr), 0)
        flags = 0
        Size = os.path.getsize(self.BINFile)
        size = Size - self.offsetAddr
        buffer = binfile.read(size)
        for i in range(0, int(size) - 3):
            if ((buffer[i+1] != b'') and ((buffer[i + 1] & 0xf8) == 0xf0)):  # 分支前缀指令1111 0 signed 11bit prefix offset poff
                if ((buffer[i+3] != b'') and ((buffer[i + 3] & 0xf8) == 0xf8)):
                    flags = 1
                    localaddr = self.offsetAddr + i
                    self.calculate_thumbBL_jumpAddr(buffer, i, localaddr)
        for j in range(0, int(size) - 3):
            if ((buffer != b'') and ((buffer[j + 1] & 0xf8) == 0xf0)):  # 分支前缀指令1111 0 signed 11bit prefix offset poff
                if ((buffer != b'') and ((buffer[j + 3] & 0xf8) == 0xe8)):
                    flags = 2
                    localaddr = self.offsetAddr + j
                    self.calculate_thumb_BLX_jumpAddr(buffer, j, localaddr)
        binfile.close()

    def write_excel(self):
        data2_excel(self.localladdr, self.dstaddr, self.offsetaddr)

    def calculate_thumbBL_jumpAddr(self, wdhex, k, loc_addr):
        local_address = loc_addr   # local_address 为j*1000 + k
        print('this is a BL branch command')
        self.localladdr.append(hex(local_address))  # 把本地地址加入本地地址列表
        print('local_address is ' + hex(local_address))
        print('assembly command  {:2x}-{:2x}-{:2x}-{:2x}'.format(wdhex[k], wdhex[k + 1], wdhex[k + 2],
                                                                 wdhex[k + 3]))
        highaddroffset = (((wdhex[k + 1] & 0x7) << 8) + (wdhex[k]))  # 取分支指令后11位poff
        lowaddroffset = (((wdhex[k + 3] & 0x7) << 8) + (wdhex[k + 2]))  # 取低字节后11位
        hiaddo = highaddroffset << 12  # poff<<12
        loaddo = lowaddroffset << 1  # offset<<2
        countoffset = hiaddo | loaddo  # 高和低相或
        if (((countoffset & 0x7fffff) >> 22) == 1):  # 最高位符号位为1，代表向前跳转，需要-1然后取反
            print('jump to back')
            actaddr = 0xffffffff & (~(countoffset - 1))  # -1后取反
            deactaddr = actaddr & 0x7fffff  # 取低23位地址
            dstaddress = local_address + 0x4 - deactaddr  # 目的地址: instruction+4+ (poff<<12) + offset*2
            if (dstaddress < 0):
                dstaddress001 = ~-dstaddress + 1 & 0xffffffff  # 如果是负值则取补码
                print('the destination address is ' + hex(dstaddress001))
                offsetaddress = dstaddress001 - local_address  # 得到偏移地址
                print('offset' + hex(offsetaddress))
                self.dstaddr.append(hex(dstaddress001))  # 加入目的地址列表
            else:
                offsetaddress = abs(dstaddress - local_address)  # 得到偏移地址
                self.dstaddr.append(hex(dstaddress))  # 加入目的地址列表

            self.offsetaddr.append(hex(offsetaddress))  # 加入偏移地址列表

        elif (((countoffset & 0x7fffff) >> 22) == 0):  # 如果第23位为0 则直接计算目标地址
            print('jump to front')
            dstaddress = local_address + countoffset + 4
            print('the destination address is ' + hex(dstaddress))
            offsetaddress1 = abs(local_address - dstaddress)
            print('offset' + hex(offsetaddress1))
            self.dstaddr.append(hex(dstaddress))
            self.offsetaddr.append(hex(offsetaddress1))
    def calculate_thumb_BLX_jumpAddr(self,wdbin3,n,loc_addr3):
        localaddr3 = loc_addr3 + n
        self.localladdr.append(hex(localaddr3))
        print('This is a Thumb BLX command')
        print('Local address is ' + hex(localaddr3))
        # blx指令跳转目标地址为：(instruction + 4 + (poff<<12) + offset*4) & ~3
        countoffset1 = (((wdbin3[n + 1] & 0x7) << 8) | (wdbin3[n + 0])) << 12
        countoffset2 = ((wdbin3[n + 3] & 0x7) << 7) | ((wdbin3[n + 2] & 0xfe) >> 1)
        countoffadd = countoffset1 | (countoffset2 << 2)
        if ((countoffadd & 0x7fffff) >> 22 == 1):  # 最高位符号位为1，代表向前跳转，需要-1然后取反
            print('jump to front')
            actaddr3 = 0xffffffff & (~(countoffadd - 1))  # -1后取反
            print('actaddr3 address is ' + hex(actaddr3))
            deactaddr3 = actaddr3 & 0x7fffff  # 取低23位地址
            dstaddress3 = (localaddr3 + 0x4 - deactaddr3) & ~3  # 得到目的地址
            if (dstaddress3 < 0):
                dstaddress10 = ~-dstaddress3 + 1 & 0xffffffff  # 取补码扩展到32位
                print('the destination address is ' + hex(dstaddress10))
                offsetaddress3 = abs(dstaddress10 - localaddr3)  # 得到偏移地址
                print('offset' + hex(offsetaddress3))
                self.offsetaddr.append(hex(offsetaddress3))  # 加入偏移地址列表
                self.dstaddr.append(hex(dstaddress10))  # 加入目的地址列表
            else:
                print('the destination address is ' + hex(dstaddress3))
                offsetaddress3 = abs(dstaddress3 - localaddr3)  # 得到偏移地址
                print('offset' + hex(offsetaddress3))
                self.dstaddr.append(hex(dstaddress3))  # 加入目的地址列表
                self.offsetaddr.append(hex(offsetaddress3))  # 加入偏移地址列表
        elif ((countoffadd & 0x7fffff) >> 22 == 0):  # 如果第23位为0 则直接计算目标地址
            print('jump to back')
            dstaddr3 = (localaddr3 + 4 + countoffadd) & ~3
            print('the destination address is ' + hex(dstaddr3))
            offsetaddress3 = abs(localaddr3 - dstaddr3)
            print('offset' + hex(offsetaddress3))
            self.dstaddr.append(hex(dstaddr3))
            self.offsetaddr.append(hex(offsetaddress3))




if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Jump instruction address')

    parser.add_argument('-f', '--value2', type=str, dest='fileName')
    parser.add_argument('-d', '--hex', type=int, dest='startAddr')
    parser.add_argument('-c', '--value3', type=str, dest='commandJumpAddress')

    args = parser.parse_args()

    calc_instant = None
    #######################################################
    if args.commandJumpAddress == 'ARM':
        calc_instant = calculate_armJumpInstruction_address(args.fileName, args.startAddr)
        print('ARMJUMP')
    if args.commandJumpAddress == 'THUMB':
        calc_instant = calculate_thumbInstruction_jumpAddr(args.fileName, args.startAddr)
        print('THUMBJUMP')
    ##########################################################


    if calc_instant != None :
        calc_instant.judge_jump_instruction()
        calc_instant.write_excel()


