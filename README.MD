Создать файлы отчетов в директории reports
```bash
go run main.go
```
Запустить benchmark
```bash
go test -bench=.
```
Запустить benchmark 1 с профилированием и посмотреть через go tool pproof
```bash
GOGC=off go test -bench=BenchmarkFirstExcel -o=test1 -cpuprofile test1.out
go tool pprof -web ./test1 ./test1.out
```
Запустить benchmark 2 с профилированием и посмотреть через go tool pproof
```bash
GOGC=off go test -bench=BenchmarkSecondExcel -o=test2 -cpuprofile test2.out
go tool pprof -web ./test2 ./test2.out
```