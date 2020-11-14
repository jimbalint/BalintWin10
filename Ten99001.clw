

   MEMBER('ten99.clw')                                ! This is a MEMBER module

                     MAP
                       INCLUDE('TEN99001.INC'),ONCE        !Local module procedure declarations
                     END


Main                 PROCEDURE                        ! Declare Procedure
    map
ZipString       procedure(string),string
    end

    ! files

InFile       file,driver('ascii'),name(fName),pre(in),create,bindable,thread
record       record
x            string(255)
             .  .

OutFile      file,driver('ascii','/CLIP=off'),name(txtName),pre(out),create,bindable,thread
record       record
x            string(750)
             .  .

tpsFile    file,driver('topspeed'),name('f:\asend\HOF_2015\Ten99.tps'),pre(tps),create,bindable,thread
RECORD     RECORD,PRE()

FID         string(10)
FID2        string(9)
NameID      string(50)
Payer1      string(50)
Payer2      string(50)
Payer3      string(50)
PayerCity   string(40)
PayerState  string(2)
PayerZip    string(9)
PayerPhone  string(40)

PayeeID     string(20)
PayeeID2    string(20)
AmountLine  byte
Box         string(20)
PayeeName   string(50)
Amount      decimal(12,2)
PayeeAddr   string(50)
PayeeCity   string(21)
PayeeState  string(2)
PayeeZip    string(9)
        .   .
DirQue  queue(file:queue),pre(dq)
            end
    ! variables

bAcctNum    long

msg1        string(60)
msg2        string(60)
msg3        string(60)


TTIN        string(9)

x           string(255)

PayYear     long
CCode       string(5)
TestFile    string(1)

TName1      string(40)
TName2      string(40)
TCompName1  string(40)
TCompName2  string(40)
TCompAddr   string(40)
TCompCity   string(40)
TCompState  string(2)
TCompZip    string(9)

ContactName     string(40)
ContactPhone    string(15)
ContactEMail    string(50)

LastFID         string(10)

Total1          decimal(12,2)
Total6          decimal(12,2)
Total9          decimal(12,2)

ACount          long
BCount          long
BCountTl        long

SeqNumber       long

TotalAmount     decimal(12,2)

RunType         string(1)

ct          long
ct2         long
ct3         long
ct4         long

i           long
j           long
k           long

Trans           group,over(out:Record),pre(T)
Type            string(1)
PayYear         string(4)
PriorYear       string(1)
TIN             string(9)
CCode           string(5)
b1              string(7)
TestFile        string(1)
Foreign         string(1)
Name            string(40)
Name2           string(40)
CompName        string(40)
CompName2       string(40)
CompAddr        string(40)
CompCity        string(40)
CompState       string(2)
CompZip         string(9)
b2              string(15)
PayeeCt         string(8)
ContactName     string(40)
ContactPhone    string(15)
ContactEMail    string(50)
Tape            string(2)
MediaNum        string(6)
b3              string(83)
SeqNumber       string(8)
b4              string(10)
VendInd         string(1)
VendName        string(40)
VendAddr        string(40)
VendCity        string(40)
VendState       string(2)
VendZip         string(9)
VendContact     string(40)
VendPhone       string(15)
b5              string(35)
VendForeign     string(1)
b6              string(8)
b7              string(2)
            .

Payer           group,over(out:Record),pre(A)
Type            string(1)
PayYear         string(4)
b1              string(6)
TIN             string(9)
NameControl     string(4)
LastFiling      string(1)
! Combined        string(1)  taken out fortax year 2010
ReturnType      string(2)    ! A

! AmountCodes     string(14)   ! 3 = Other Income   4 = FWT
AmountCodes     string(16)   ! 3 = Other Income   4 = FWT  - 2011 expanded 2 positions - 2013 - added "9"
                            ! 2015 - removed "9"

! b2              string(10) - 2011 changed from 10 to 8
b2              string(8)

Foreign         string(1)
Name1           string(40)
Name2           string(40)
XferAgent       string(1)
ShipAddr        string(40)
City            string(40)
State           string(2)
Zip             string(9)
Phone           string(15)
b3              string(260)
SeqNumber       string(8)
b4              string(241)
b5              string(2)
        .

Payee           group,over(out:Record),pre(B)
Type            string(1)
PayYear         string(4)
Corrected       string(1)
NameControl     string(4)   !!!!!!!!!!!!!!!!!
TINType         string(1)   ! 1 = EIN  2 = SSN
TIN             string(9)
AcctNum         string(20)
OfficeCode      string(4)
b1              string(10)
Pay01           string(12)
Pay02           string(12)
Pay03           string(12)
Pay04           string(12)
Pay05           string(12)
Pay06           string(12)
Pay07           string(12)
Pay08           string(12)
Pay09           string(12)
Pay10           string(12)
Pay11           string(12)
Pay12           string(12)
Pay13           string(12)
Pay14           string(12)
b2              string(24)
Foreign         string(1)
Name1           string(40)
Name2           string(40)
b3              string(40)
Addr1           string(40)
b4              string(40)
City            string(40)
State           string(2)
Zip             string(9)
b5              string(1)
SeqNumber       string(8)
b6              string(36)
TIN2            string(1)
b7              string(2)
DirectSales     string(1)
b8              string(115)
SpecialData     string(60)
SWT             string(12)
CWT             string(12)
CombCode        string(2)
b9              string(2)
            .

PayerEnd        group,over(out:record),pre(C)
Type            string(1)
Count           string(8)
b1              string(6)
Total01         string(18)
Total02         string(18)
Total03         string(18)
Total04         string(18)
Total05         string(18)
Total06         string(18)
Total07         string(18)
Total08         string(18)
Total09         string(18)
Total10         string(18)
Total11         string(18)
Total12         string(18)
Total13         string(18)
Total14         string(18)
b2              string(232)
SeqNumber       string(8)
b3              string(241)
b4              string(2)
            .


Final           group,over(out:Record),pre(F)
Type            string(1)
CountA          string(8)
Zero            string(21)
b1              string(19)
PayeeCt         string(8)
b2              string(442)
SeqNumber       string(8)
b3              string(241)
b4              string(2)
            .



Window WINDOW('Caption'),AT(,,311,136),GRAY
       STRING(@s60),AT(37,18,233,10),USE(msg1),CENTER
       STRING(@s60),AT(38,40,233,10),USE(msg2),CENTER
       STRING(@s60),AT(36,64,233,10),USE(msg3),CENTER
     END
  CODE
    ! main code

    open(Window)

    !! *************************************************
    !!
    !! ** HEY **
    !! use TopScan to convert TPS to TXT
    !! shortcut is on desktop
    !!  --> column / select all / export ....
    !!
    !! *************************************************

    !!!!!!!!!!!!!!!!!!!!!!!!!!!!

    ! >>>> for tax year 2009
    !      "B" record Account number
    !      use incremental counter or
    !      some kind of unique number
    !      DO NOT USE THE TIN !!!
    !
    !     >>> changed to b:AcctNum - record b pos 21 - 40
    ! >>>>

    PayYear         = 2015

    ! ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    !!! TestFile        = 'T'
    TestFile        = ''
    ! ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    ContactName     = 'Rebecca Foldi'
    ContactPhone    = '330-849-6926'
    ContactEMail    = 'RFoldi@INVENT.ORG'

    RunType = 'b'

    bAcctNum = 0

    ! tax year 2010 notes
    !     https://fire.irs.gov
    !     rfoldi / Runner11
    !
    ! use CCode 17677 for both files
    ! use 341580038 / 17677 to upload both files
    !
    !
    ! **** check TestFile string above ****

    ! tax year 2011 notes
    ! password changed to Runner12
    !
    ! Payer (type A) record changed

    if RunType = 'a'

!        FName = 'c:\asend\hof_2010\KIDSMISC1099-2012IRS.txt'         ! IN
!        txtName = 'c:\asend\hof_2010\bKIDSMISC1099-2012IRS.txt'         ! OUT
!        CCode = '17G22'
!        CCode = '17677'         ! changed to same as run "b" for tax year 2010
!
!        TTIN = '521088781'     ! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< was 341766051 in 2009
!        TName1      = 'Invent Now Kids, Inc.'
!        TName2      = ''
!        TCompName1  = 'Invent Now Kids, Inc.'
!        TCompName2  = ''
!        TCompAddr   = '221 S. Broadway'
!        TCompCity   = 'Akron'
!        TCompState  = 'OH'
!        TCompZip    = '44308'

    elsif RunType = 'b'             ! only 'b' was done for 2011 2012 2013

        ! IN
        FName = 'f:\asend\hof_2015\NIHFI1099-2015-FINAL-IRSversion032316.txt'

        ! OUT
        txtName = 'f:\asend\hof_2015\NIHFI1099-2015-FINAL-IRSversion032316_a.txt'

        CCode = '17677'

        TTIN = '341580038'
        TName1      = 'National Inventors Hall of Fame Foundation, Inc.'
        TName2      = ''
        TCompName1  = 'National Inventors Hall of Fame Foundation, Inc.'
        TCompName2  = ''
        TCompAddr   = '221 S. Broadway'
        TCompCity   = 'Akron'
        TCompState  = 'OH'
        TCompZip    = '44308'

!    elsif RunType = 'c'         ! correction for 2008
!
!        FName = 'c:\asend\hof\INOWCorr.TXT'          ! IN
!        txtName = 'c:\asend\hof\hofCORR99.txt'          ! OUT
!        CCode = '17677'
!
!        TTIN = '341580038'
!        TName1      = 'National Inventors Hall of Fame Foundation, Inc.'
!        TName2      = ''
!        TCompName1  = 'National Inventors Hall of Fame Foundation, Inc.'
!        TCompName2  = ''
!        TCompAddr   = '221 S. Broadway'
!        TCompCity   = 'Akron'
!        TCompState  = 'OH'
!        TCompZip    = '44308'

    else
        stop('Run Type ? ')
    end

    !!!!!!!!!!!!!!!!!!!!!!!!!!!!

    open(InFile)
    if error() then halt(0,'InFile open: ' & error()).

    create(OutFile)
    if error() then halt(0,'OutFile Create: ' & error()).
    open(OutFile)
    if error() then halt(0,'OutFile Open: ' & error()).

    create(tpsFile)
    if error() then halt(0,'tps Create: ' & error()).
    open(tpsFile,12h)
    if error() then halt(0,'tps Open: ' & error()).
    stream(tpsFile)

    set(InFile)
    loop
        next(InFile)
        if error() then break.

        i += 1
        if i = 15
            j += 1
            i = 1
            msg2 = '1099s processed: ' & j
            display()

            TotalAmount += tps:Amount

            ! output
            ! check data
            tps:NameID = clip(tps:FID) & ' ' & tps:Payer1
            append(tpsFile)
            clear(tps:Record)

        end

        case i

            of 1
                tps:Payer1 = sub(in:x,6,37)
                if clip(sub(in:x,43,10)) <> ''
                    tps:Amount = sub(in:x,43,10)
                    tps:AmountLine = 1
                    tps:Box = 'Box #1 Rents'
                end

            of 2
                tps:Payer2 = sub(in:x,6,50)

            of 3
                tps:Payer3 = sub(in:x,6,50)

            of 4
                tps:PayerCity    = sub(in:x,6,20)
                tps:PayerState   = sub(in:x,27,2)
                tps:PayerZip     = ZipString(sub(in:x,30,10))

            of 5
                tps:PayerPhone = sub(in:x,6,50)

            of 6
                if clip(sub(in:x,43,10)) <> ''
                    tps:Amount = sub(in:x,43,10)
                    tps:AmountLine = 6
                    tps:Box = 'Box #3 Other Income'
                end

            of 7    ! SS#
                tps:FID = sub(in:x,6,10)
                loop k = 1 to len(tps:FID)
                    if InString(sub(tps:FID,k,1),'0123456789',1,1)
                        tps:FID2 = clip(tps:FID2) & sub(tps:FID,k,1)
                    end
                end

                tps:PayeeID = sub(in:x,23,20)

                tps:PayeeID2 = ''
                loop k = 1 to len(tps:PayeeID)
                    if InString(sub(tps:PayeeID,k,1),'0123456789',1,1)
                        tps:PayeeID2 = clip(tps:PayeeID2) & sub(tps:PayeeID,k,1)
                    end
                end

            of 8    ! name
                tps:PayeeName = sub(in:x,6,50)

            of 9    ! amount
                if clip(sub(in:x,43,10)) <> ''
                    tps:Amount = sub(in:x,43,10)
                    tps:AmountLine = 9
                    tps:Box = 'Box #7 Non Emp Comp'
                end

            of 11   ! address
                tps:PayeeAddr = sub(in:x,6,50)

            of 12   ! csz
                tps:PayeeCity = sub(in:x,6,21)
                tps:PayeeState = upper(sub(in:x,27,2))
                tps:PayeeZip = ZipString(sub(in:x,30,10))

        end

    end

    j += 1
    i = 1
    msg2 = '1099s processed: ' & j
    display()

    TotalAmount += tps:Amount

    ! output
    ! check data
    tps:NameID = clip(tps:FID) & ' ' & tps:Payer1
    append(tpsFile)
    clear(tps:Record)

    flush(tpsFile)
    build(tpsFile)

    ct = 0

    !!!!
    ContactPhone = tps:PayerPhone
    LastFID = ''

    set(tpsFile)
    loop
        next(tpsFile)
        if error() then break.

        ct += 1
        msg3 = format(ct,@n5) & ' 1099s Written'
        display()

        if ct = 1

            SeqNumber = 1

            ! T Record
            t:Type          = 'T'
            t:PayYear       = PayYear
            t:PriorYear     = ''
            t:TIN           = tps:FID2           ! ?????
            t:CCode         = CCode
            t:TestFile      = TestFile
            t:Foreign       = ''
            t:Name          = TName1
            t:Name2         = TName2
            t:CompName      = TCompName1
            t:CompName2     = TCompName2
            t:CompAddr      = TCompAddr
            t:CompCity      = TCompCity
            t:CompState     = TCompState
            t:CompZip       = TCompZip
            t:PayeeCt       = format(j,@n08)
            t:ContactName   = ContactName
            t:ContactPhone  = ContactPhone
            t:ContactEMail  = ContactEMail
            t:Tape          = ''
            t:MediaNum      = ''
            t:SeqNumber     = format(SeqNumber,@n08)
            t:VendInd       = 'I'
            t:VendName      = ''
            t:VendAddr      = ''
            t:VendCity      = ''
            t:VendState     = ''
            t:VendZip       = ''
            t:VendContact   = ''
            t:VendPhone     = ''
            t:VendForeign   = ''

            add(OutFile)
            SeqNumber += 1

        end

        if LastFID <> tps:FID

            if LastFID <> ''

                ! Type C record
                clear(out:Record)
                c:Type          = 'C'
                c:Count         = format(BCount,@n08)

                c:Total01       = format(Total1 * 100,@n018)
                c:Total02       = all('0')
                c:Total03       = format(Total6 * 100,@n018)
                c:Total04       = all('0')
                c:Total05       = all('0')
                c:Total06       = all('0')
                c:Total07       = format(Total9 * 100,@n018)
                c:Total08       = all('0')
                c:Total09       = all('0')
                c:Total10       = all('0')
                c:Total11       = all('0')
                c:Total12       = all('0')
                c:Total13       = all('0')
                c:Total14       = all('0')

                c:SeqNumber     = format(SeqNumber,@n08)

                add(OutFile)
                SeqNumber += 1

                Total1 = 0
                Total6 = 0
                Total9 = 0
                BCount = 0

            end

            ! A Record
            clear(out:Record)
            a:Type          = 'A'
            a:PayYear       = PayYear
            a:TIN           = tps:FID2           ! ?????????
            a:NameControl   = ''
            a:LastFiling    = ''
            ! a:Combined      = ''
            a:ReturnType    = 'A '               !
            a:AmountCodes   = '12345678ABCDE'
            a:Foreign       = ''
            a:Name1         = tps:Payer1
            a:Name2         = tps:Payer2
            a:XferAgent     = '0'
            a:ShipAddr      = tps:Payer3
            a:City          = tps:PayerCity
            a:State         = tps:PayerState
            a:Zip           = tps:PayerZip
            a:Phone         = tps:PayerPhone
            a:SeqNumber     = format(SeqNumber,@n08)

            add(OutFile)
            SeqNumber += 1
            ACount += 1

        end
        LastFID = tps:FID

        clear(out:Record)
        b:Type          = 'B'
        b:PayYear       = PayYear

        if RunType = 'c'
            b:Corrected     = 'G'
        else
            b:Corrected     = ''
        end

        b:NameControl   = ''            ! ?

        if tps:PayeeID = '' or len(tps:PayeeID) < 3
            b:TINType = ''
        elsif sub(tps:PayeeID,3,1) = '-'        ! EIN
            b:TINType = '1'
        elsif sub(tps:PayeeID,4,1) = '-'        ! SSN
            b:TINType = '2'
        else
            b:TINType = ''
        end

        b:TIN           = tps:PayeeID2

        ! changed in tax year 2009
        !   use unique number
        bAcctNum += 1
        b:AcctNum       = format(bAcctNum,@n020)

        b:OfficeCode    = ''

        b:Pay01         = all('0',12)
        b:Pay02         = all('0',12)
        b:Pay03         = all('0',12)
        b:Pay04         = all('0',12)
        b:Pay05         = all('0',12)
        b:Pay06         = all('0',12)
        b:Pay07         = all('0',12)
        b:Pay08         = all('0',12)
        b:Pay09         = all('0',12)
        b:Pay10         = all('0',12)
        b:Pay11         = all('0',12)
        b:Pay12         = all('0',12)
        b:Pay13         = all('0',12)
        b:Pay14         = all('0',12)

        if tps:AmountLine = 1       ! Box #1 Rents
            b:Pay01 = format(tps:Amount * 100,@n012)
            Total1 += tps:Amount
        elsif tps:AmountLine = 6    ! Box #3 Other Income
            b:Pay03 = format(tps:Amount * 100,@n012)
            Total6 += tps:Amount
        elsif tps:AmountLine = 9    ! Box #7 Non Emp Comp
            b:Pay07 = format(tps:Amount * 100,@n012)
            Total9 += tps:Amount
        else
            ! 20150324
            ! stop('Amount Line ? ' & tps:AmountLine)
        end

        b:Foreign       = ''
        b:Name1         = tps:PayeeName
        b:Name2         = ''
        b:Addr1         = tps:PayeeAddr
        b:City          = tps:PayeeCity
        b:State         = tps:PayeeState
        b:Zip           = ZipString(tps:PayeeZip)
        b:SeqNumber     = format(SeqNumber,@n08)
        b:TIN2          = ''
        b:DirectSales   = ''
        b:SpecialData   = ''
        b:SWT           = ''
        b:CWT           = ''
        b:CombCode      = ''

        add(OutFile)
        SeqNumber += 1
        BCount += 1
        BCountTl += 1

    end

    ! Type C record
    clear(out:Record)
    c:Type          = 'C'
    c:Count         = format(BCount,@n08)

    c:Total01       = format(Total1 * 100,@n018)
    c:Total02       = all('0')
    c:Total03       = format(Total6 * 100,@n018)
    c:Total04       = all('0')
    c:Total05       = all('0')
    c:Total06       = all('0')
    c:Total07       = format(Total9 * 100,@n018)
    c:Total08       = all('0')
    c:Total09       = all('0')
    c:Total10       = all('0')
    c:Total11       = all('0')
    c:Total12       = all('0')
    c:Total13       = all('0')
    c:Total14       = all('0')

    c:SeqNumber     = format(SeqNumber,@n08)

    add(OutFile)
    SeqNumber += 1

    ! type F record
    clear(out:Record)
    f:Type          = 'F'
    f:CountA        = format(ACount,@n08)
    f:Zero          = all('0')
    f:PayeeCt       = format(BCountTl,@n08)
    f:SeqNumber     = format(SeqNumber,@n08)
    add(OutFile)

    message('1099s processed: ' & format(j,@n6) & ' ' & format(TotalAmount,@n14.2))



ZipString       procedure(InString)

z1      byte
z2      byte
ReturnString    string(9)


    code

    if len(clip(InString)) <= 5
        return(InString)
    end

    loop z1 = 1 to len(InString)
        if sub(InString,z1,1) = '-' then cycle.
        ReturnString = clip(ReturnString) & sub(InString,z1,1)
    end

    return(ReturnString)

