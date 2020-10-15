package com.wdh.exceldemo.test;


import java.util.Scanner;

public class AnswerTwo {
    public static class Node {
        // 优先级
        private int[] p;

        private Node next;

        public int[] getP() {
            return p;
        }

        public void setP(int[] p) {
            this.p = p;
        }

        public Node getNext() {
            return next;
        }

        public void setNext(Node next) {
            this.next = next;
        }

        public Node(int[] p) {
            this.p = p;
            this.next = null;
        }

        public Node() {

        }
    }

    public static class pQueue {
        private Node head;

        private Node tail;

        private int minP;

        public Node getHead() {
            return head;
        }

        public void setHead(Node head) {
            this.head = head;
        }

        public Node getTail() {
            return tail;
        }

        public void setTail(Node tail) {
            this.tail = tail;
        }

        public int getMinP() {
            return minP;
        }

        public void setMinP(int minP) {
            this.minP = minP;
        }

        public pQueue() {
            if (null == head) {
                head = new Node();
                tail = head;
                head.setNext(null);
                minP = 0;
            }
        }

        public void add(Node node) {
            if (null == head) {
            } else if (null == head.getNext()) {
                head.setNext(node);
                tail = node;
                minP = node.getP()[1];
            } else if (node.getP()[1] <= minP) {
                if (node.getP()[0] == tail.getP()[0] && node.getP()[1] == minP) {
                } else {
                    tail.setNext(node);
                    tail = node;
                    minP = node.getP()[1];
                }
            } else {
                Node curr = head.getNext();
                Node pre = head;
                while (curr != null && curr.getP()[1] >= node.getP()[1]) {
                    pre = curr;
                    curr = curr.getNext();
                }
                if (curr == null) {
                    if (node.getP()[0] == pre.getP()[0] && node.getP()[1] == pre.getP()[1]) {
                    } else {
                        pre.setNext(node);
                        tail = node;
                        minP = node.getP()[1];
                    }
                } else {
                    if (pre.getP() != null && node.getP()[0] == pre.getP()[0] && node.getP()[1] == pre.getP()[1]) {
                    } else {
                        node.setNext(curr);
                        pre.setNext(node);
                    }
                }
            }
        }

        public void print() {
            Node node = head.getNext();
            StringBuffer sb = new StringBuffer();
            while (null != node) {
                sb.append(node.getP()[0]).append(",");
                node = node.getNext();
            }
            System.out.println(sb.toString().substring(0, sb.toString().length() - 1));
        }
    }

    public static void main(String[] args) {

        Scanner in = new Scanner(System.in);
        String str = in.nextLine();

        String newStr = str.replace("(","").replace(")","");
        String[] arr = newStr.split(",");
        pQueue queue = new pQueue();
        for (int i = 0; i < arr.length; i++) {
            if (i % 2 == 0) {
                int[] p= {Integer.valueOf(arr[i]),Integer.valueOf(arr[i+1])};
                queue.add(new Node(p));
            }
        }

        queue.print();
    }
}
