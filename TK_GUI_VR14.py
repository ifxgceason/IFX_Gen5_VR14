import tkinter as tk
from tkinter import messagebox as mb
import IFX_VR14_3D
import IFX_VR14_DC_LL

def func_3d():
    

    rail_name=variable.get()
    start_freq=float(input_text2.get())
    end_freq=float(input_text3.get())
    icc_min=int(input_text4.get())
    icc_max=int(input_text5.get())
    freq_steps_per_decade=int(input_text6.get())
    rise_time=int(input_text7.get())
    cool_down_delay=int(input_text8.get())
    start_duty=int(input_text9.get())
    end_duty=int(input_text10.get())
    duty_step=int(input_text11.get())
    vout=eval(input_text12.get())
    excel=bool(input_checkbox_14.get())
    IFX_VR14_3D.vr14_3d(rail_name,vout,start_freq,end_freq,icc_min,icc_max,freq_steps_per_decade,rise_time,cool_down_delay,start_duty,end_duty,duty_step)

        
def dc_loadline():
    rail_name=variable.get()
    start_freq=float(input_text2.get())
    end_freq=float(input_text3.get())
    icc_min=int(input_text4.get())
    icc_max=int(input_text5.get())
    freq_steps_per_decade=int(input_text6.get())
    rise_time=int(input_text7.get())
    cool_down_delay=int(input_text8.get())
    start_duty=int(input_text9.get())
    end_duty=int(input_text10.get())
    duty_step=int(input_text11.get())
    vout=eval(input_text12.get())
    LL_point=eval(input_text13.get())
    excel=bool(input_checkbox_14.get())
    IFX_VR14_DC_LL.vr14_ifx_dc(rail_name,vout,icc_max,cool_down_delay,LL_point,excel)
    




def about_message():
    mb.showinfo("About", "IFX VR14 GUI 2020/08/21", detail="https://github.com/ifxgceason/Gen5_VR14")

     
app = tk.Tk()
app.title("IFX VR14 GUI")
app.geometry('550x550')

#fix window size. 
app.resizable(0, 0)

# Adding Menu and commands
tk.menubar = tk.Menu(app)
file = tk.Menu(tk.menubar, tearoff = 0) 
tk.menubar.add_cascade(label ='Help', menu = file) 

file.add_separator()
file.add_command(label ='about', command = about_message)

# set GUI column and row index 
i=0
j=0

#set label 1
label_1 = tk.Label(app,text = "rail_name")
label_1.grid(column=i, row=j, sticky=tk.W)

#set label 2
label_2 = tk.Label(app,text = "start_freq(KHz)")
label_2.grid(column=i, row=j+1, sticky=tk.W)

#set label 3
label_3 = tk.Label(app,text = "end_freq(KHz")
label_3.grid(column=i, row=j+2, sticky=tk.W)

#set label 4
label_4 = tk.Label(app,text = "icc_min(A)")
label_4.grid(column=i, row=j+3, sticky=tk.W)

#set label 5
label_5 = tk.Label(app,text = "icc_max(A)")
label_5.grid(column=i, row=j+4, sticky=tk.W)

#set label 6
label_6 = tk.Label(app,text = "freq_steps_per_decade")
label_6.grid(column=i, row=j+5, sticky=tk.W)

#set label 7
label_7 = tk.Label(app,text = "rise_time(nSec)")
label_7.grid(column=i, row=j+6, sticky=tk.W)

#set label 8
label_8 = tk.Label(app,text = "cool_down_delay(Sec)")
label_8.grid(column=i, row=j+7, sticky=tk.W)

#set label 9
label_9 = tk.Label(app,text = "start_duty(%)")
label_9.grid(column=i, row=j+8, sticky=tk.W)

#set label 10
label_10 = tk.Label(app,text = "end_duty(%)")
label_10.grid(column=i, row=j+9, sticky=tk.W)

#set label 11
label_11 = tk.Label(app,text = "duty_step(%)")
label_11.grid(column=i, row=j+10, sticky=tk.W)

#set label 12
label_12 = tk.Label(app,text = "vout(V)")
label_12.grid(column=i, row=j+11, sticky=tk.W)

#set label 13
label_13 = tk.Label(app,text = "LoadLine Iout list")
label_13.grid(column=i, row=j+12, sticky=tk.W)

#set label 14
##label_14 = tk.Label(app,text = "select Excel")
##label_14.grid(column=i, row=j+13, sticky=tk.W)

#set label 15
label_15 = tk.Label(app,text = " ")
label_15.grid(column=i, row=j+14, sticky=tk.W)


#set input text box 1
variable=tk.StringVar(app)
variable.set("VCCIN")
optionmenu=tk.OptionMenu(app,variable,"VCCIN","VCCINFAON","VCCFA_EHV","VCCFA_EHV_FIVRA","VCCD_HV")
optionmenu.grid(column=i+1,row=j)



#set input test box 2
input_text2 = tk.StringVar(value="0.3")
entrystartPS = tk.Entry(app, width=10, textvariable=input_text2)
entrystartPS.grid(column=i+1, row=j+1, padx=20)

#set input text box 3
input_text3 = tk.StringVar(value="100")
entryVout = tk.Entry(app, width=10, textvariable=input_text3)
entryVout.grid(column=i+1, row=j+2, padx=20)

#set input text box 4
input_text4 = tk.StringVar(value="6")
entryIcc = tk.Entry(app, width=10, textvariable=input_text4)
entryIcc.grid(column=i+1, row=j+3, padx=20)

#set input text box 5
input_text5 = tk.StringVar(value="405")
entryIoutStep = tk.Entry(app, width=10, textvariable=input_text5)
entryIoutStep.grid(column=i+1, row=j+4, padx=20)

#set input text box 6
input_text6 = tk.StringVar(value="10")
entryDelayTime = tk.Entry(app, width=10, textvariable=input_text6)
entryDelayTime.grid(column=i+1, row=j+5, padx=20)

#set input text box 7
input_text7 = tk.StringVar(value="315")
entryVstart = tk.Entry(app, width=10, textvariable=input_text7)
entryVstart.grid(column=i+1, row=j+6, padx=20)

#set input test box 8
input_text8 = tk.StringVar(value="2")
entrystartPS = tk.Entry(app, width=10, textvariable=input_text8)
entrystartPS.grid(column=i+1, row=j+7, padx=20)

#set input text box 9
input_text9 = tk.StringVar(value="10")
entryVout = tk.Entry(app, width=10, textvariable=input_text9)
entryVout.grid(column=i+1, row=j+8, padx=20)

#set input text box 10
input_text10 = tk.StringVar(value="50")
entryIcc = tk.Entry(app, width=10, textvariable=input_text10)
entryIcc.grid(column=i+1, row=j+9, padx=20)

#set input text box 11
input_text11 = tk.StringVar(value="10")
entryIoutStep = tk.Entry(app, width=10, textvariable=input_text11)
entryIoutStep.grid(column=i+1, row=j+10, padx=20)

#set input text box 12
input_text12 = tk.StringVar(value="[1.83,1.73]")
entryIoutStep = tk.Entry(app, width=10, textvariable=input_text12)
entryIoutStep.grid(column=i+1, row=j+11, padx=20)

#set input text box 13
input_text13 = tk.StringVar(value="[0,20,40,60,80,100,120,140,160,180,200,220,240,260,280,300,320,240,360,380,400]")
entryIoutStep = tk.Entry(app, width=70, textvariable=input_text13)
entryIoutStep.grid(column=i, row=j+13, padx=20,columnspan=3,sticky=tk.E+tk.W)

#set input text box 14
input_checkbox_14 = tk.IntVar()
c_14=tk.Checkbutton(app, text='Use Excel to open result', onvalue=1, offvalue=0,variable=input_checkbox_14)
c_14.grid(column=i, row=j+14, padx=10)
c_14.select()

#set input text box 15
#input_text15 = tk.StringVar(value=" ")
#entryIoutStep = tk.Entry(app, width=10, textvariable=input_text15)
#entryIoutStep.grid(column=i+1, row=j+14, padx=20)

# set logo
img_png = tk.PhotoImage(file = 'ifx.dll')
label_img = tk.Label(app, image = img_png)
label_img.grid(column=0,row=16)

#set run button1

resultButton = tk.Button(app, text = '3D test',height = 2, width = 15,command=func_3d)
resultButton.grid(column=0, row=17,columnspan=2, pady=10, sticky=tk.E+tk.W)

#set run button2

resultButton = tk.Button(app, text = 'DC LoadLine test',height = 2, width = 30,command=dc_loadline)
resultButton.grid(column=0, row=18,columnspan=2, pady=10, sticky=tk.E+tk.W)

# set result 
resultString=tk.StringVar()
resultLabel = tk.Label(app, textvariable=resultString,fg="black")
resultLabel.grid(column=1, row=19, padx=8, sticky=tk.W)

#set meanbar

def main():
    app.config(menu=tk.menubar)
    app.mainloop()

if __name__ == "__main__":
    main()

