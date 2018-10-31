# json_to_excel

Convert data.json to out.xlsx 
Nested object will be object_assingned for the data to be single level.

Example :

```
{
  title: 'some title'
  user: {
    name: 'creator name'
  }
}
```

Will be leveled to header 

"title", "user.name"
