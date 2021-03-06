# README
 一个读取 Excel 内容并使用云片接口进行群发短信的程序

## 程序说明
### config.json
|字段|说明|备注|
|-|-|-|
|apikey|调用接口所需要的 `APIKEY` |在后台可以查询到|
|excel-file|所读入的 excel 文件||
|log-file|程序日志所写入的文件|若不存在会自动创建|
|template|包含使用的模板的信息|具体见下方介绍，该对象中必含字段为 `content`|
|template>content|模板的内容|必须是短信全文（用于云片API的智能匹配），变量用 `%s` 表示，变量将会在发送前使用 `python` 的字符串展开功能替换之。|
|var-count|模板中所包含的变量个数|十分重要，读取 excel 文件中的多少列由该设置决定，且进行字符串展开时这个配置也起到了相当的作用，设置不正确可能提前报错退出。|
|log|关于日志的相关设置|对象|
|log>disabled|是否禁用日志|值为 `true` 或 `false`|
|log>level|记录的日志等级|该程序中等级从 `1` 到 `5`，设置的值越高，记录的信息越详细|

## 使用时的注意事项
 * 读取表格的时候第一列会作为电话号码列读入，之后的列要读入多少由 `config.json` 中的 `var-count` 字段决定，同时这也对应着调用短信接口时作为参数的那些字段数量。
 * 在短信模板中使用 `%s` 来代表变量，届时将会使用 `python` 的展开字符串功能去使用变量替换。
 * `Excel` 文件中所有字段必须为 `文本` 格式，否则将可能引发不可预料的错误，所以建议在确认发出短信前先查看列表中有哪些数据。
 * `var-count` 一定要设置正确，否则会在展开字符串时抛出错误。
