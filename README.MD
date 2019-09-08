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