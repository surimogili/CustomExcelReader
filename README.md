# CustomExcelReader
Read excel and convert to object

Today we have so many examples to read excel and convert into object. But we are facing diffculty when converting object has one to many realtion with other objects like below

public class Employee
{
String name;
String email;
List<Address> addresses;
}

public class Address
{
String city;
String state;
String pincode;
}

This code will help you to achieve above requirement
