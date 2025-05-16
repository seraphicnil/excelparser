package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
)

//TIP To run your code, right-click the code and select <b>Run</b>. Alternatively, click
// the <icon src="AllIcons.Actions.Execute"/> icon in the gutter and select the <b>Run</b> menu item from here.

// ./app -s=./source_dir -e=./export_dir
func main() {
	flag.StringVar(&SourceDir, "s", "source", "excel文件目录")
	flag.StringVar(&ExportDir, "e", "export", "导出目录")
	flag.Parse()

	fmt.Println("SourceDir", SourceDir)
	fmt.Println("ExportDir", ExportDir)
	err := filepath.WalkDir(SourceDir, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			fmt.Printf("访问路径出错 %q: %v\n", path, err)
			if Interrupt {
				return err
			}
			return nil // 忽略错误，继续遍历
		}
		err = ParseOnFile(path)
		if err != nil {
			if Interrupt {
				return err
			}
		}
		return nil
	})

	if err != nil {
		fmt.Printf("遍历目录时出错: %v\n", err)
	}
}

//TIP See GoLand help at <a href="https://www.jetbrains.com/help/go/">jetbrains.com/help/go/</a>.
// Also, you can try interactive lessons for GoLand by selecting 'Help | Learn IDE Features' from the main menu.
