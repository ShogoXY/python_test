thisdict =	{
  "brand": ["Ford","opel","mazda"],
  "model": ["Mustang","astra","6"],
  "kraj": ["polska","niemcy","wlochy"]
}

value1="a"
while value1!="":
    value1=input("key\n")

    while value1 in thisdict:
        value=input("value\n")

        if value in thisdict[value1]:
            print(f"Yes, Value: '{value}' exists in dictionary")
        else:
            print(f"No, Value: '{value}' does not exists in dictionary")
        if value=="":
            break    
        
print ("KONIEC")
