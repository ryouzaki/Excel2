import pymongo
from pymongo import MongoClient


def main():
    client = MongoClient('localhost', 27017)
    db = client['VizeumHealth']
    collection = db['Placements']
    db=mongoDB()
    print (db.updateCreative("AK_VN_005_001"))
    #print(collection.find_one({"creatives_list": {"creative_id":"AK_VN_005_001"}}))
    #collection.update_one({PLACEMENT_ID:placement_id}, {"$set": {field:value}}, upsert=True)





if __name__ == '__main__':
    main()
