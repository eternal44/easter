# When working with excel files gets too repetitive...
This easter egg is for you.

## Installation
```
$ /usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"
$ brew install git
$ brew cask install libreoffice
$ brew install node
$ npm install
```

## Usage
Follow these steps:
- Label your file as `target.xlsx`
- Place it on the project root

and execute the following command:
```
$ npm start
```

You'll find your sorted & converted PDFs in a directory called
`convertedPDFFiles` and all excel files (except the original
`target.xlsx`) in the sortedExcelFiles.

## Credits
To my one and only.  We pair well together.

## License
MIT

## To do
- ES6 conversion
- async file writes and reads
- hande errors
