
print ("Program sprawdzający dane w dwóch listach")
list_1=( "a", "b", "c", "d", "e", )
list_2=( "1", "2", "3", "4", "5", )
list_3=( "11", "22", "33", "44", "55", )
list_4=( "111", "222", "333", "444", "555", )
list_5=( "1111", "2222", "3333", "4444", "5555", )
list_6=( "11111", "22222", "33333", "44444", "55555", )
list=(list_2,list_3,list_4,list_5,list_6)
w_1=input("podaj wartość 1\n")
w_id=list_1.index(w_1)
print (w_id)

while w_1 in list_1:
    
    if w_1 in list_1 :

        print("lista dla "+ w_1 +"\n")
        print (list[w_id])
        w_2=input("\npodaj wartość 2\n")
        if w_2 in list[w_id]:
            print("zgadza się")
        else:
            print("nie ma takiej wartości")
    else:
        print("podaj inna wartość")

    w_1=input("podaj wartość 1\n")
    w_id=list_1.index(w_1)
print ("błędna wartość, koniec programu")
