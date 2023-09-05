

def get_code_name(my_list:list,value:str):
    for i,a in enumerate(my_list) :
        if a.startswith(value):
            return my_list[i]