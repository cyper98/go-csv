package go_csv

import (
	"archive/zip"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
)

var (
	ErrEmptyPath = errors.New("zip: empty path list")
)

type zipEntry struct {
	name string
	data []byte
}

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

func ZipBuffersParallel(outputZipPath string, files map[string][]byte) error {
	if len(files) == 0 {
		return ErrEmptyPath
	}

	entries := make(chan zipEntry, len(files))
	for name, data := range files {
		entries <- zipEntry{name: name, data: data}
	}
	close(entries)

	f, err := os.Create(outputZipPath)
	if err != nil {
		return err
	}
	defer f.Close()

	zw := zip.NewWriter(f)
	for e := range entries {
		w, err := zw.Create(slashPath(e.name))
		if err != nil {
			_ = zw.Close()
			return err
		}
		if _, err := w.Write(e.data); err != nil {
			_ = zw.Close()
			return err
		}
	}

	return zw.Close()
}

func ZipPathsParallel(outputZipPath string, paths []string, baseDir string) error {
	if len(paths) == 0 {
		return ErrEmptyPath
	}

	entries := make(chan zipEntry, len(paths))
	for _, p := range paths {
		data, err := os.ReadFile(p)
		if err != nil {
			continue
		}
		rel := p
		if baseDir != "" {
			if r, err := filepath.Rel(baseDir, p); err == nil {
				rel = r
			}
		}
		entries <- zipEntry{name: rel, data: data}
	}
	close(entries)

	f, err := os.Create(outputZipPath)
	if err != nil {
		return err
	}
	defer f.Close()

	zw := zip.NewWriter(f)
	for e := range entries {
		w, err := zw.Create(slashPath(e.name))
		if err != nil {
			_ = zw.Close()
			return err
		}
		if _, err := w.Write(e.data); err != nil {
			_ = zw.Close()
			return err
		}
	}

	return zw.Close()
}

func slashPath(p string) string {
	return strings.ReplaceAll(p, string(filepath.Separator), "/")
}
