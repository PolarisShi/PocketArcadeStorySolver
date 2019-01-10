import pandas as pd
import numpy as np

def main(num):
    
    data = pd.read_excel('data.xlsx', sheet_name='Sheet1',usecols=[1,2,3])
    score = pd.read_excel('data.xlsx', sheet_name='Sheet1',usecols=[4,5])
    
    print(data)
    
    data = np.array(data)
    score = np.array(score)
    data_list=data.tolist()
    
    l = len(data_list)
    for ii in range(l):
        if "*" in data_list[ii]:
            data_list[ii].remove("*")
    
    name = list(set(data[:,0].tolist()+data[:,1].tolist()+data[:,2].tolist()))
    name.remove("*")
    
    n = len(name)
    
    def contain(test_list, data_list, score):
        count = 0
        a_sum = 0
        c_sum = 0
        l = len(data_list)
        for itr in range(l):
            set_data = set(data_list[itr])
            list_set_data = list(set_data)
            flag = 0
            for ii in range(len(list_set_data)):
                if test_list.count(list_set_data[ii]) < data_list[itr].count(list_set_data[ii]):
                    flag = 1
                    break
            if flag == 0:
                a_sum += score[itr][0]
                c_sum += score[itr][1]
                count += 1
        return count, a_sum, c_sum
    
    L = []
    L_count = []
    L_a = []
    L_c = []
    for i1 in range(n):
        for i2 in range(i1,n):
                test_list = [name[i1], name[i2]]
                test_list.sort()
                count, a_sum, c_sum = contain(test_list, data_list, score)
                if count > 0:
                    L.append(test_list)
                    L_count.append(count)
                    L_a.append(a_sum)
                    L_c.append(c_sum)
    
    limit_list = [2, 3, 5, 6, 8, 10, 12, 14, 16, 18]
    import copy
    for kk in range(num-2):
        L_new = []
        L_count_new = []
        L_a = []
        L_c = []
        for i1 in range(len(L)):
            for i2 in range(n):
                test_list = L[i1] +[name[i2]]
                test_list.sort()
                count, a_sum, c_sum = contain(test_list, data_list, score)
                if count >= limit_list[kk] and test_list not in L_new:
                    L_new.append(test_list)
                    L_count_new.append(count)
                    L_a.append(a_sum)
                    L_c.append(c_sum)
                    
        L = copy.deepcopy(L_new)
        L_count = copy.deepcopy(L_count_new)
        
    return L, L_count, L_a, L_c

num = 12
L, L_count, L_a, L_c = main(num)

print("Areas:", num)
best_ii = 0
best_count = max(L_count)
print("Max Themed Areas:", best_count)
for ii in range(len(L)):
    if L_count[ii] == best_count and L_c[ii] > L_c[best_ii]:
        best_ii = ii

best_L = L[best_ii]
best_a = L_a[best_ii]
best_c = L_c[best_ii]
print(best_L)
print("Appeal:", best_a, "  Cost:", best_c)
