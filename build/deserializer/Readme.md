# Deserializer

## Overview

**deserializer** allow you to automatically get deserialized stacktrace in errors

## How to use

* Put files with errors in json format in any `logs-dir`

* Make `unique-file` with unique errors that `last-log-file` adds(`last-log-file` is optional)

    ```bash
    node parse-json-log-dir.js logs-dir last-log-file unique-file
    ```
* Install dependencies for `download-maps.js`

    ```bash
    npm ci
    ```

* Download Closure Compiler maps into `maps-dir` for version specified in `unique-file`

    ```bash
    node download-maps.js accessKeyId, secretAccessKey unique-file maps-dir
    ```
* Download Closure Compiler maps into `maps-dir` for version specified in `unique-file`

    ```bash
    node download-maps.js accessKeyId, secretAccessKey unique-file maps-dir
    ```

maps-dir structure will be

```
maps-dir/
    7.4.1-36
        word.props.js.map
        cell.props.js.map
        slide.props.js.map
    <version>-<build>
        <maps>
```

* Deserialize `unique-file` call stack into `deserialized-file` (docx or txt by extension)

    ```bash
    node deserialize.js unique-file deserialized-file maps-dir
    ```
  
* The result will be in the `deserialized-file`

