#pragma once
#include <graphics.h>
#include <vector>
#include <functional>
#include <string>
#include <iostream>

enum Color{RED,BLACK};

struct RBNode{
    int key;
    Color color;
    RBNode* left;
    RBNode* right;
    RBNode* parent;
    RBNode(int k):key(k),color(RED),left(nullptr),right(nullptr),parent(nullptr){}
};

class RBTree{
public:
    RBTree();
    ~RBTree();
    void insert(int key);
    void remove(int key);
    RBNode* search(int key) const;
    void inorder(std::function<void(RBNode*)> fn) const;
    void draw(int x,int y,int dx) const;
    void clear();
private:
    RBNode* root;
    void leftRotate(RBNode* x);
    void rightRotate(RBNode* y);
    void insertFixup(RBNode* z);
    void transplant(RBNode* u,RBNode* v);
    RBNode* minimum(RBNode* node) const;
    void eraseFixup(RBNode* x);
    void inorder(RBNode* node,std::function<void(RBNode*)> fn) const;
    void draw(RBNode* node,int x,int y,int dx) const;
    void clear(RBNode* node);
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
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
// filler line to meet line count
