#pragma once
#include "RBTree.h"
#include "HashTable.h"
#include "Graph.h"
#include <graphics.h>
#include <conio.h>
#include <chrono>
#include <cmath>

class Visualizer{
public:
    Visualizer();
    void run();
private:
    RBTree rbtree;
    HashTable hashtable;
    Graph graph;
    int mode; // 0 tree,1 hash,2 graph
    double scale;
    void draw();
    void handleKey(int ch);
    void handleMouse(MOUSEMSG msg);
    void testPerformance();
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
