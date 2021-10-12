import webbrowser
import re
# dobra dodatkowy test
print("program podaje lokalizację w mapach google")
ulica = input("podaj nazwę ulicy:\n")
numer = input("podaj numer domu:\n")
code = input("Podaj kod pocztowy" +
             "\nKod pocztowy musi składać się z samych liczb: \nnp: 12345 \n")
while True:

    if re.match(r"[0-9-]+", code) and len(code) > 4 and len(code) < 7 != None:
        code = re.sub("[ ()-]", '', code)  # remove space, (), -
        # code2=(code[:2]+"-"+code[3:6])
        code2 = (f"{code[:2]}-{code[2:5]}")
        print(code2)

        break

    else:
        print("zły kod pocztowy")
        print("kod poczatowy musi miec skladać się z samych liczb")
        code = input("Podaj kod pocztowy: ")
        continue

miasto = input("podaj miasto:\n")

webbrowser.open('https://www.google.pl/maps/place/' +
                ulica+"+"+numer+"+"+code2+"+"+miasto)
