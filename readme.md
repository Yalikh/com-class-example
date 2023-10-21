# Пример использования в VisualFoxPro 9 COM-объекта, созданного под .NET Framework 4.8

## Реализация объекта, важные моменты

* Необходимо маркировать класс одним из атрибутов `ComVisibleAttribute` или `ClassInterfaceAttribute`
* Необходимо маркировать класс атрибутом `GuidAttribute` - он задаёт ID компонента

``` C#
[ComVisible(true)]
[Guid("07bf25e6-ad4c-4bf7-890c-213c7f5d01cf")]
[ClassInterface(ClassInterfaceType.None)]
public class MyComClass
...
```

* Нужно включить настройку `Register for COM Interop` в свойствах проекта. Если этого не сделать, то при регистрации компонента нужно указывать ключ `/tlb`, чтобы tlb-файл всё-таки создался.

## Регистрация объекта

Регистрация
``` Powershell
RegAsm.exe ".\MyComLib\bin\Debug\MyComLib.dll" /codebase /tlb
```
*Ключ `/codebase` обязятелен для добавления сборки в GAC.*

&nbsp;

Отмена регистрации
``` Powershell
RegAsm.exe ".\MyComLib\bin\Debug\MyComLib.dll" /unregister
```

Проверить зарегистрирован ли компонент
``` Powershell

# Поиск по ID
Get-CimInstance Win32_ClassicCOMClassSetting | Where-Object { $_.ComponentId -eq "{07bf25e6-ad4c-4bf7-890c-213c7f5d01cf}" }

# Происк по имени класса
Get-CimInstance Win32_ClassicCOMClassSetting | Where-Object { $_.ProgId -eq "MyComLib.MyComClass" }
```

## Ссылки
* [Пример создания COM-класса](https://learn.microsoft.com/en-us/dotnet/csharp/advanced-topics/interop/example-com-class)
* [Советы по регистрации компонента](https://stackoverflow.com/questions/7092553/turn-a-simple-c-sharp-dll-into-a-com-interop-component)
* [Описание утилиты RegAsm](https://learn.microsoft.com/ru-ru/dotnet/framework/tools/regasm-exe-assembly-registration-tool)
* [Советы по просмотру списка COM-классов (Powershell 5)](https://stackoverflow.com/questions/660319/where-can-i-find-all-of-the-com-objects-that-can-be-created-in-powershell)
* [Советы по просмотру списка COM-классов (Powershell 7)](https://forums.powershell.org/t/powershell-7-is-missing-get-wmiobject/14011)
