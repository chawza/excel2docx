# Excel2docx
my helper scripts to convert manual testcase to document while working at BTPN

# How to Install

1. install python from [python foundation](https://www.python.org/downloads/). My test enviornment uses version 3.7, I have not test higher version (try it if you want)
2. open your terminal
3. install excel2docx from this github repo <b>globaly</b>
    ```
    python -m pip install git+https://github.com/chawza/excel2docx.git@main
    ```
    > some machine uses ```python2``` or ```python3```

    > you can change the ```main``` option to other branch if you want
    

# How to run
## The Gui way
open the app using graphical user interface for easier experience thanks to Tkinter module
!(Excel2docx in GUI version)[/docs/imgs/gui-preview.png]
1. open your terminal
2. run the package as python module
   ```
   python -m excel2docx gui
   ```
3. the gui should show

   > Make sure the python have [Tkinter](https://docs.python.org/3/library/tkinter.html) module installed

   > You can create cmd/bash shortcut to run this quickly

## The CLI way
quick way for converting files or including in your scripts
1. open terminal
2. run the package as python module
   ```
    python -m excel2docx [source file] [target directory] [--uac-tc]
   ```
   > source file: full path of your excel file in your machine
   >
   > target directory: (<b>optional</b>) where do you want your docx file saved in your machine
   >
   > --uac-tc: if the scenario title is in the testcase sheet (status: about to drop)
   
