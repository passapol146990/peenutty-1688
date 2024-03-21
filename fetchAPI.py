import requests,json
uriAPI = 'http://localhost:8080/'
def fetchAPI(selectGroup):
    uri = f"{uriAPI}getdatabybot1688?key=382557d3731870efb6e863c174ccb351f191987e79a9c5805e16698e055b4364&&group={selectGroup}"
    try:
        return requests.get(uri);
    except:
        return 404
# # len(json.loads(fetchAPI("อุปกรณ์-อิเล็กทรอนิกส์").text))
# res = json.loads(fetchAPI("บ้านและไลฟ์สไตล์").text)
# print(len(res))
def FetchAPIweb(selectGroup):
    uri = f"{uriAPI}getdatas?group={selectGroup}"
    try:
        return requests.get(uri);
    except:
        return 404
res = json.loads(FetchAPIweb("บ้านและไลฟ์สไตล์").text)
print(len(res))



# df = open("data.txt", "a",encoding="utf8")
# for i in range(1,100):
#     for j in range(2,20):
#         # print("{a} ^ {b} \t= {c:.5f}".format(a=i/100,b=j,c=(i/100)**j))
#         df.write("{a} ^ {b} \t= {c:.8f}\n".fo
# rmat(a=i/100,b=j,c=(i/100)**j))
#         # print("%^% = %.2f" % i/10,j,(i/10)**j)
#     df.write("//////////////////////////////////////////////////////\n")
#     # print("//////////////////////////////////////////////////////")
# e = 2.71828
# for i in range(-20,20):
#     df.write("{a}^{b} = {c}\n".format(a=e,b=i,c=e**i))
#     # print("{a}^{b} = {c}".format(a=e,b=i,c=e**i))