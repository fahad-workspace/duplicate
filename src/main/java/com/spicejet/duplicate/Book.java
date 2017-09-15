package com.spicejet.duplicate;

public class Book {

    private String first;
    private String second;

    public static Book getCopiedBookInstance(Book book) {
        Book value = new Book();
        value.setFirst(book.getFirst());
        value.setSecond(book.getSecond());
        return value;
    }

    public String getFirst() {
        return first;
    }

    public void setFirst(String first) {
        this.first = first;
    }

    public String getSecond() {
        return second;
    }

    public void setSecond(String second) {
        this.second = second;
    }
}
