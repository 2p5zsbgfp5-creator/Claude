#!/usr/bin/env python3
"""Independent Python replication of the model's core math to validate logic
and expected liquidity path (Base scenario)."""
W=[0.070,0.070,0.085,0.085,0.090,0.085,0.070,0.060,0.090,0.100,0.105,0.090]
OMZ=5_000_000; MRG=0.28; BTW=0.21; DPM=30.4
DSO,DPO,DIO=45,40,55; GR=0.0; MA=0.0; KG=0.0
deb0,vrd0,crd0,bank0=600000,750000,500000,300000
loon,wgl,vak,wi=40000,0.28,0.08,0.0
loan0,afl_m,rente=900000,10000,0.06
cats=[("V",8000,0,"JA"),("V",2800,0,"JA"),("P",0,0.020,"JA"),("V",3500,0,"JA"),
      ("V",2000,0,"NEE"),("P",0,0.010,"JA"),("V",2500,0,"JA"),("V",2000,0,"JA"),("V",3000,0,"JA")]
invs=[(4,45000),(9,80000),(14,60000),(20,40000)]
def run(gr=GR, ma=MA, dso=DSO, dpo=DPO, dio=DIO, kg=KG, wi=wi, extra_capex=0, capexmnd=12, label="Base"):
    bank=bank0; prev_deb,prev_vrd,prev_crd,prev_loan=deb0,vrd0,crd0,loan0
    mins=1e18; maxtekort=0; ends=[]
    for k in range(1,25):
        m=((k-1)%12)
        rev=W[m]*OMZ*(1+gr)
        marg=MRG+ma
        cogs=rev*(1-marg)
        vrd=dio/DPM*cogs
        purch=cogs+vrd-prev_vrd
        purchi=purch*(1+BTW)
        revincl=rev*(1+BTW)
        deb=dso/DPM*revincl
        crd=dpo/DPM*purchi
        receipts=prev_deb+revincl-deb
        suppay=prev_crd+purchi-crd
        # personnel
        yr=(k-1)//12
        vakadd=loon*12*vak if m==4 else 0
        pers=(loon*(1+wgl)+vakadd)*(1+wi)*(1+kg*yr)
        # other costs
        other=0; other_btw=0
        for typ,vast,pct,btw in cats:
            val=vast*(1+kg*yr) if typ=="V" else pct*rev
            other+=val
            if btw=="JA": other_btw+=val
        # capex
        capex=sum(b for mm,b in invs if mm==k)+(extra_capex if k==capexmnd else 0)
        # financing
        afl=min(afl_m,prev_loan)
        loan=prev_loan-afl
        rente_c=prev_loan*rente/12
        # VAT
        vat_out=rev*BTW
        vat_in=(purch+other_btw+capex)*BTW
        # store accruals for quarterly pay
        if k==1: accr=[]
        accr.append(vat_out-vat_in)
        # quarterly: pay in k=4,7,10,... sum prev 3
        vat_pay=0
        if k>=4 and (k-1)%3==0:
            vat_pay=max(0,sum(accr[k-4:k-1]))
        # taxes (loonheffing zit al in personeelskas-uit; NIET apart meetellen)
        vpb=8000 if k%3==0 else 0
        # prive
        prive=40000 if k in (6,18) else 4000
        # liquidity
        totin=receipts+0+0
        totuit=suppay+pers+other+vat_pay+vpb+capex+rente_c+afl+prive+0
        net=totin-totuit
        bank=bank+net
        ends.append(bank)
        mins=min(mins,bank); maxtekort=max(maxtekort,max(0,100000-bank))
        prev_deb,prev_vrd,prev_crd,prev_loan=deb,vrd,crd,loan
    print(f"{label:22s} min={mins:>12,.0f}  maxFinBeh={maxtekort:>10,.0f}  end24={ends[-1]:>12,.0f}")
    return ends,mins,maxtekort

print("=== Cross-check core (buffer 100k) ===")
b,_,_=run(label="Base")
run(gr=0.10, label="Omzet +10% (Best-omzet)")
run(gr=-0.10, label="Omzet -10%")
run(ma=-0.03, label="Marge -3ppt")
run(dso=60, label="DSO 60 (+15)")
run(dio=75, label="DIO 75 (+20)")
run(dpo=30, label="DPO 30 (-10)")
run(wi=0.10, label="Loon +10%")
run(extra_capex=500000, capexmnd=6, label="Extra capex 500k m6")
run(gr=-0.10, ma=-0.03, dso=60, dio=75, dpo=30, label="Worst combi")
print("\nBase eindliquiditeit per maand:")
print("  " + "  ".join(f"{i+1}:{v/1000:.0f}k" for i,v in enumerate(b)))
