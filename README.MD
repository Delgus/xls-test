### Cравнение производительности библиотек github.com/360EntSecGroup-Skylar/excelize/v2 и github.com/tealeg/xlsx

Бенчмарки Ubuntu 18
```
goos: linux
goarch: amd64
pkg: github.com/delgus/turnover-test
BenchmarkExcelizeLibLow-2                     30          53943161 ns/op         5414533 B/op      57842 allocs/op
BenchmarkXLSXLibLow-2                         30          41464037 ns/op         3552118 B/op      48887 allocs/op
BenchmarkExcelizeLibHigh-2                     1        2106927137 ns/op        122156200 B/op   3923061 allocs/op
BenchmarkXLSXLibHigh-2                         1        1806870655 ns/op        207134184 B/op   2041891 allocs/op
BenchmarkExcelizeLibHighParallel-2             1        5164669614 ns/op        757775864 B/op  11170673 allocs/op
BenchmarkXLSXLibHighParallel-2                 1        2511535619 ns/op        414266368 B/op   4083771 allocs/op
BenchmarkExcelizeLibFile-2                     1        4240230552 ns/op        378351032 B/op   5585370 allocs/op
BenchmarkXLSXLibFile-2                         1        1709010318 ns/op        206105928 B/op   2041886 allocs/op
PASS
ok      github.com/delgus/xlsx-test 20.616s

```

Включить модули если необходимо
```bash
export GO111MODULE=on
```

Установить Graphviz для визуального отображения результатов бенчмарков
```bash
sudo apt install graphviz
```

Создать файлы отчетов в директории reports
```bash
go run main.go
```
Запустить benchmark
```bash
go test -bench=. -benchmem
```
Запустить benchmark с профилированием cpu и памяти
```bash
GOGC=off go test -bench=. -benchmem -o=test -cpuprofile cpu.out -memprofile=mem.out
```
Посмотреть через go tool pprof
```bash
//cpu
go tool pprof -web ./test ./cpu.out
//memory
go tool pprof -web ./test ./mem.out
```
