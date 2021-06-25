import pandas as pd
import pathlib
from pathlib import Path
import difflib
import time



#file_path = "\\rubus\Data_trumph\Tatra\BAZZA\DETALI"
file_path = "."

# In[2]:


df = pd.read_excel("4.1_Сводный_план_1го_участка_18.06.21.xlsx")


# In[3]:


df.columns = df.iloc[4]
df


# In[4]:


diap_1 = int(input("Введите номер верхней границы: "))
diap_2 = int(input("Введите номер нижней границы: ")) + 1


# In[5]:


df  = df.iloc[diap_1:diap_2]


# In[6]:


df = df.groupby('Обозначение детали').size() #здесь  больше не трогать

print("В выбранном диапазоне существуют следующие детали:\n")
print(df, "\n")

# df = pd.Series({'ANR.00.01.038': 29, 'ANR.00.01.037': 18, 'BNR.00.01.039': 37})


# In[13]:


start_time = time.time()


# In[8]:


def find_a_number(dir):
    number = ""
    for i in str(dir)[::-1]:
        if  i != "_":
            try:
                number += str(int(i))
            except:
                continue
        else:
            break
    return int(number[::-1])


# In[9]:


def list_part_nums(name_of_detail, dir): #зависим от find_a_number()
    list_of_details = []
    for i in dir: # цикл, который ищет подстроку детали в списке всех деталей и добавляет имя файла в массив
        if name_of_detail in str(i).replace("." , "").replace("_" , "").replace("-" , "").replace("/", ""): # test -- works!!!
            list_of_details.append(find_a_number(str(i)))
    return sorted(list_of_details)[::-1]


def list_of_links(name_of_detail, dir, target_number): #зависим от find_a_number()
    list_of_links = {}
    for i in dir: # цикл, который ищет подстроку детали в списке всех деталей и добавляет имя файла в массив
        if name_of_detail in str(i).replace("." , "").replace("_" , "").replace("-" , "").replace("/", ""): # test -- works!!!
            list_of_links.update({find_a_number(str(i)):i})
    return list_of_links.get(target_number)


# In[12]:


#df1[df1.index[3]]
#df1["ANR.08.L1.000.002"]
#list_of_details = list_part_nums(name_of_detail, dir) если что разкомментить
#number = 0 #нужна для того чтобы переменная target_number была видимой
def return_res(name_of_detail_orig):
    for target_number in list_of_details: #цикл, который ищет подходящий файл на основе сопоставления цифр
        if df[name_of_detail_orig] >= target_number: #!!!поменять строку на переменную детали df to df
            break
    url = str(list_of_links(name_of_detail, dir, target_number))    
    return {"Насчитанно " + name_of_detail_orig:df[name_of_detail_orig], "Будет выполнено ":df[name_of_detail_orig]//target_number * target_number, "Остаток":df[name_of_detail_orig] - df[name_of_detail_orig]//target_number * target_number, "Ссылка":url}


dir = list(Path(file_path).rglob("*.html" ))


for i in range(len(df.index)): #rename df to df
    name_of_detail_orig = df.index[i] #rename df to df
    name_of_detail = name_of_detail_orig.replace("." , "").replace(" нерж" , "").replace(" лев" , "").replace(" прав" , "")
    
    list_of_details = list_part_nums(name_of_detail, dir)
    if list_of_details != []:
        print(return_res(name_of_detail_orig)) # засунуть в pandas frame чтобы потом переконвертировать в таблицу   

        
print("Время выполнения: %s с. " % round((time.time() - start_time), 2))


