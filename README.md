Read xlsx to json
===

## Feature
* read and parse excel to json object
* sheet ref to array supported
* file callback
* embed ojbect by xx.xx

## Types
* `int64` to go int64
* `float64` to go float64
* `string` to go string
* `time` to go time.Time
* `file` to call OnFile, default as string value

## Example
### Excel
* sheet:`product`
| id,int64    | user_id,int64 | name,string | no,string | price,float64 |  create_time,time  |
| ----------- | ------------- | ----------- | --------- | ------------- | ------------------ |
| 1           | 1             | product01   | 10001     | 100           | 2022/6/9  13:04:05 |
| 2           | 2             | product02   | 10002     | 100           | 2022/6/9  13:04:05 |
* sheet:`user`
| id,int64    | username,string | password,string | avatar,file | money,float64 | register_time,time | company.name,string | company.address:o,string | products,ref,product,user_id=id |
| ----------- | --------------- | --------------- | ----------- | ------------- | ------------------ | ------------------- | ------------------------ | ------------------------------- |
| 1           | user01          | 123             | user01.png  | 100.51        | 2022/6/9  13:04:05 | company             | address                  |                                 |
| 2           | user02          | 123             | user01.png  | 0             | 2022/6/9  13:04:05 | company             |                          |                                 |

### Code
```.go
users, err := reader.Read("user0", 1, 0)
if err != nil {
    t.Error(err)
    return
}
fmt.Println(converter.JSON(users))
```

### Result
```.json
[
    {
        "avatar": "user01.png",
        "company": {
            "address": "address",
            "name": "company"
        },
        "id": 1,
        "money": 100.51,
        "password": "123",
        "products": [
            {
                "create_time": "2022-06-09T13:04:04.999999855+08:00",
                "id": 1,
                "name": "product01",
                "no": "10001",
                "price": 100,
                "user_id": 1
            }
        ],
        "register_time": "2022-06-07T11:04:05.000000274+08:00",
        "username": "user01"
    },
    {
        "avatar": "user02.png",
        "company": {
            "name": "company"
        },
        "id": 2,
        "money": 0,
        "password": "123",
        "products": [
            {
                "create_time": "2022-06-09T13:04:04.999999855+08:00",
                "id": 2,
                "name": "product02",
                "no": "10002",
                "price": 100,
                "user_id": 2
            }
        ],
        "register_time": "2022-06-07T11:04:05.000000274+08:00",
        "username": "user02"
    }
]
```