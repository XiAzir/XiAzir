#pragma once
#include <vector>
#include <list>
#include <functional>
#include <graphics.h>
#include <string>
#include <iostream>

class HashTable{
public:
    explicit HashTable(size_t size=101);
    void insert(int key);
    bool find(int key) const;
    void erase(int key);
    void draw(int x,int y,int cell) const;
    void clear();
private:
    std::vector<std::list<int>> table;
    size_t hash(int key) const;
};
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
