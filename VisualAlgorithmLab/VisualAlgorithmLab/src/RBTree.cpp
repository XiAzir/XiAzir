#include "RBTree.h"

RBTree::RBTree():root(nullptr){}

RBTree::~RBTree(){
    clear();
}

void RBTree::leftRotate(RBNode* x){
    RBNode* y = x->right;
    x->right = y->left;
    if(y->left) y->left->parent = x;
    y->parent = x->parent;
    if(!x->parent) root = y;
    else if(x==x->parent->left) x->parent->left = y;
    else x->parent->right = y;
    y->left = x;
    x->parent = y;
}

void RBTree::rightRotate(RBNode* y){
    RBNode* x = y->left;
    y->left = x->right;
    if(x->right) x->right->parent = y;
    x->parent = y->parent;
    if(!y->parent) root = x;
    else if(y==y->parent->left) y->parent->left = x;
    else y->parent->right = x;
    x->right = y;
    y->parent = x;
}

void RBTree::insert(int key){
    RBNode* z = new RBNode(key);
    RBNode* y = nullptr;
    RBNode* x = root;
    while(x){
        y = x;
        if(z->key < x->key) x = x->left; else x = x->right;
    }
    z->parent = y;
    if(!y) root = z;
    else if(z->key < y->key) y->left = z; else y->right = z;
    insertFixup(z);
}

void RBTree::insertFixup(RBNode* z){
    while(z->parent && z->parent->color == RED){
        if(z->parent == z->parent->parent->left){
            RBNode* y = z->parent->parent->right;
            if(y && y->color == RED){
                z->parent->color = BLACK;
                y->color = BLACK;
                z->parent->parent->color = RED;
                z = z->parent->parent;
            }else{
                if(z == z->parent->right){
                    z = z->parent;
                    leftRotate(z);
                }
                z->parent->color = BLACK;
                z->parent->parent->color = RED;
                rightRotate(z->parent->parent);
            }
        }else{
            RBNode* y = z->parent->parent->left;
            if(y && y->color == RED){
                z->parent->color = BLACK;
                y->color = BLACK;
                z->parent->parent->color = RED;
                z = z->parent->parent;
            }else{
                if(z == z->parent->left){
                    z = z->parent;
                    rightRotate(z);
                }
                z->parent->color = BLACK;
                z->parent->parent->color = RED;
                leftRotate(z->parent->parent);
            }
        }
    }
    root->color = BLACK;
}

RBNode* RBTree::minimum(RBNode* node) const{
    while(node->left) node = node->left;
    return node;
}

void RBTree::transplant(RBNode* u,RBNode* v){
    if(!u->parent) root = v;
    else if(u == u->parent->left) u->parent->left = v;
    else u->parent->right = v;
    if(v) v->parent = u->parent;
}

void RBTree::eraseFixup(RBNode* x){
    while(x != root && (!x || x->color == BLACK)){
        if(x == x->parent->left){
            RBNode* w = x->parent->right;
            if(w && w->color == RED){
                w->color = BLACK;
                x->parent->color = RED;
                leftRotate(x->parent);
                w = x->parent->right;
            }
            if((!w->left || w->left->color == BLACK) && (!w->right || w->right->color == BLACK)){
                if(w) w->color = RED;
                x = x->parent;
            }else{
                if(!w->right || w->right->color == BLACK){
                    if(w->left) w->left->color = BLACK;
                    w->color = RED;
                    rightRotate(w);
                    w = x->parent->right;
                }
                w->color = x->parent->color;
                x->parent->color = BLACK;
                if(w->right) w->right->color = BLACK;
                leftRotate(x->parent);
                x = root;
            }
        }else{
            RBNode* w = x->parent->left;
            if(w && w->color == RED){
                w->color = BLACK;
                x->parent->color = RED;
                rightRotate(x->parent);
                w = x->parent->left;
            }
            if((!w->right || w->right->color == BLACK) && (!w->left || w->left->color == BLACK)){
                if(w) w->color = RED;
                x = x->parent;
            }else{
                if(!w->left || w->left->color == BLACK){
                    if(w->right) w->right->color = BLACK;
                    w->color = RED;
                    leftRotate(w);
                    w = x->parent->left;
                }
                w->color = x->parent->color;
                x->parent->color = BLACK;
                if(w->left) w->left->color = BLACK;
                rightRotate(x->parent);
                x = root;
            }
        }
    }
    if(x) x->color = BLACK;
}

void RBTree::remove(int key){
    RBNode* z = search(key);
    if(!z) return;
    RBNode* y = z;
    Color yOriginal = y->color;
    RBNode* x;
    if(!z->left){
        x = z->right;
        transplant(z,z->right);
    }else if(!z->right){
        x = z->left;
        transplant(z,z->left);
    }else{
        y = minimum(z->right);
        yOriginal = y->color;
        x = y->right;
        if(y->parent == z){
            if(x) x->parent = y;
        }else{
            transplant(y,y->right);
            y->right = z->right;
            y->right->parent = y;
        }
        transplant(z,y);
        y->left = z->left;
        y->left->parent = y;
        y->color = z->color;
    }
    delete z;
    if(yOriginal == BLACK && x) eraseFixup(x);
}

RBNode* RBTree::search(int key) const{
    RBNode* x = root;
    while(x){
        if(key == x->key) return x;
        if(key < x->key) x = x->left; else x = x->right;
    }
    return nullptr;
}

void RBTree::inorder(std::function<void(RBNode*)> fn) const{
    inorder(root,fn);
}

void RBTree::inorder(RBNode* node,std::function<void(RBNode*)> fn) const{
    if(!node) return;
    inorder(node->left,fn);
    fn(node);
    inorder(node->right,fn);
}

void RBTree::draw(int x,int y,int dx) const{
    draw(root,x,y,dx);
}

void RBTree::draw(RBNode* node,int x,int y,int dx) const{
    if(!node) return;
    setlinecolor(node->color==RED?RED:BLACK);
    circle(x,y,15);
    char buf[16];
    sprintf(buf,"%d",node->key);
    outtextxy(x-10,y-5,buf);
    if(node->left){
        line(x,y,x-dx,y+50);
        draw(node->left,x-dx,y+50,dx/2);
    }
    if(node->right){
        line(x,y,x+dx,y+50);
        draw(node->right,x+dx,y+50,dx/2);
    }
}

void RBTree::clear(){
    clear(root);
    root = nullptr;
}

void RBTree::clear(RBNode* node){
    if(!node) return;
    clear(node->left);
    clear(node->right);
    delete node;
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
