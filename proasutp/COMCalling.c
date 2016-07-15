/*
    (с) 2014 www.proasutp.com

    Функция - пример работы с COM-объектами через Global Script C.

    В качестве подопытного COM-объекта выбран CCHMIRuntime, описание которого
    присутствует во встроенной помощи и который является частью WinCC, поэтому
    гарантированно присутствует на компьютере.

    Вход:
                    Нет.
            Выход:
                    Нет.
*/

void COMCalling ()
{

    const char HMIRuntimeProgID [] = "CCHMIRuntime.HMIRuntime";

    double dblValue = 0.0;

    __object * pHMIRuntimeObject = NULL;
    __object * pHMITagInterface = NULL;

    // 1. Создаем COM-объект с помощью специальной встроенной функции WinCC
    //      в качестве входного параметра используется ProgID. На выходе получаем указатель на созданный объект

    pHMIRuntimeObject = __object_create (HMIRuntimeProgID);
   
    // 2. Проверяем что объект успешно создан, если указатель равено NULL, значит объект не создан

    if (pHMIRuntimeObject == NULL) {
        printf
("Не могу создать COM-объект CCHMIRuntime.HMIRuntime\r\n");
        return;
        }

    // 3. Теперь необходимо получить указатель на COM-интерфейс работы с тегом

    pHMITagInterface = pHMIRuntimeObject->Tags ("DoubleTag");

    // 4. Проверяем что интерфейс был успешно получен

    if (pHMITagInterface == NULL) {
        printf
("Не могу получить COM-интерфейс на тег DoubleTag\r\n");
        __object_delete (pHMITagInterface);
        return;
        }

    // 5. Читаем значение тега в локальную переменную

    dblValue = pHMITagInterface->Read ();

    // 6. Проверяем что тег прочитался успешно

    if (pHMITagInterface->LastError != 0) {
        printf
("Ошибка чтения тега DoubleTag, ErrorDescription: %s\r\n", pHMITagInterface->ErrorDescription);
        }

    // 7. Делаем инкремент значения локальной переменной и записываем значение в тег

    dblValue = dblValue + 1;
    pHMITagInterface->Write (dblValue);

    // 8. Проверяем успех операции

    if (pHMITagInterface->LastError != 0) {
        printf
("Ошибка записи в тег DoubleTag, ErrorDescription: %s\r\n", pHMITagInterface->ErrorDescription);
        }

    // 9. Удаляем объекты в обратной последовательности

    __object_delete (pHMITagInterface);
    __object_delete (pHMIRuntimeObject);
}