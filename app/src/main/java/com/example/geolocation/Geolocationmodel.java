package com.example.geolocation;

import java.util.List;

public class Geolocationmodel {

    private  List<String> anyvaluerow;
    private double latitude;
    private double longitude;
    private String address;

    public List<String> getAnyvaluerow() {
        return anyvaluerow;
    }

    public void setAnyvaluerow(List<String> anyvaluerow) {
        this.anyvaluerow = anyvaluerow;
    }

    public double getLatitude() {
        return latitude;
    }

    public void setLatitude(double latitude) {
        this.latitude = latitude;
    }

    public double getLongitude() {
        return longitude;
    }

    public void setLongitude(double longitude) {
        this.longitude = longitude;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public Geolocationmodel(List<String> anyvaluerow, double latitude, double longitude, String address) {
        this.anyvaluerow = anyvaluerow;
        this.latitude = latitude;
        this.longitude = longitude;
        this.address = address;
    }

}