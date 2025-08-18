package go_csv

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// compress buffers into 1 zip
func ZipBuffers(outputZipPath string, files map[string][]byte) error {
	f, err := os.Create(outputZipPath)
	if err != nil {
		return err
	}
	defer f.Close()

	zw := zip.NewWriter(f)
	for name, data := range files {
		w, err := zw.Create(slashPath(name))
		if err != nil {
			_ = zw.Close()
			return err
		}
		if _, err := w.Write(data); err != nil {
			_ = zw.Close()
			return err
		}
	}
	return zw.Close()
}

// compress the list of files already on the disk
func ZipPaths(outputZipPath string, paths []string, baseDir string) error {
	f, err := os.Create(outputZipPath)
	if err != nil {
		return err
	}
	defer f.Close()

	zw := zip.NewWriter(f)
	for _, p := range paths {
		err := func() error {
			fp, err := os.Open(p)
			if err != nil {
				return err
			}
			defer fp.Close()

			rel := p
			if baseDir != "" {
				if r, err := filepath.Rel(baseDir, p); err == nil {
					rel = r
				}
			}
			w, err := zw.Create(slashPath(rel))
			if err != nil {
				return err
			}
			_, err = io.Copy(w, fp)
			return err
		}()
		if err != nil {
			_ = zw.Close()
			return err
		}
	}
	return zw.Close()
}

func slashPath(p string) string {
	return strings.ReplaceAll(p, string(filepath.Separator), "/")
}
