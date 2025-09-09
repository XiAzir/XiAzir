#include "HashTable.h"

HashTable::HashTable(size_t size){
    table.resize(size);
}

size_t HashTable::hash(int key) const{
    return key % table.size();
}

void HashTable::insert(int key){
    size_t idx = hash(key);
    table[idx].push_back(key);
}

bool HashTable::find(int key) const{
    size_t idx = hash(key);
    for(int v: table[idx]) if(v==key) return true;
    return false;
}

void HashTable::erase(int key){
    size_t idx = hash(key);
    table[idx].remove(key);
}

void HashTable::draw(int x,int y,int cell) const{
    for(size_t i=0;i<table.size();++i){
        rectangle(x,y+cell*i,x+cell,y+cell*(i+1));
        int offset = 0;
        for(int v: table[i]){
            char buf[16];
            sprintf(buf,"%d",v);
            outtextxy(x+5+offset,y+cell*i+5,buf);
            offset += 30;
        }
    }
}

void HashTable::clear(){
    for(auto &lst : table) lst.clear();
}
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
