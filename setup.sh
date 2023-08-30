#!/bin/zsh
#
#
#

installDir="$HOME/Documents/bakery"

red="$(tput setaf 9)"
yellow="$(tput setaf 11)"
purp="$(tput setaf 12)"
rst="$(tput sgr0)"

green=$(tput setaf 71)
blue=$(tput setaf 33)

echo "${red}            /'.                "
echo "             '. '.              "
echo "            .'|'. '.            "
echo "          .' .'  '. '.          "
echo "        .' .'      './          "
echo "   ${green} ._ ${red} \.' ${blue}      .'.         Insights Bakery Preprocessor "
echo "${green}    : :      ${blue}     : :         Setup Script"
echo "${green}    | ; _    ${blue}     ; :        "
echo "${green}    ; :'. '.   ${blue}   . :  .'\    Version 0.1"
echo "${green}    : |  '. '.  ${blue}   ''.' .'   "
echo "${green}     '.    '. '. ${blue}  .' .'     "
echo "${green}             './ ${blue}.' .'       "
echo "${green}               ${blue}.'_.'         "

read -sk 1 "? Anykey to start"

echo ""
echo "--${red} Create install directory ${rst}--"
mkdir -p "$installDir"
res="$?"
if [[ "$res" != 0 ]]; then
	echo "Something went wrong."
	exit $res
fi
cd "$installDir"
echo "--${green} Create install directory ✅ ${rst}--"

echo ""
echo "--${red} Check for Python ${rst}--"
echo "Attempting to get python version. If you do not have python installed you will get a prompt to install it"
printf "Installed Python Version: "
res="$(python3 --version)"
echo "$res"
if [[ "$res" == "" ]]; then
	echo "Python3 is not installed. Click the accept button for MacOS to install it, then re-run this script"
	exit 0
fi
echo "--${green} Check for Python ✅ ${rst}--"

echo ""

echo "--${red} Install python modules  ${rst}--"
python3 -m pip install pandas rich openpyxl requests PyInquirer
r="$?"
if [[ "$r" != "0" ]]; then
	echo "${red}Something went wrong with the pip install command${rst}"
	exit $r
fi
echo "--${green} Install python modules ✅${rst}--"

echo ""

echo "--${red} Download latest preprocessor script ${rst}--"
curl https://raw.githubusercontent.com/will-plutoflume/insights-preprocessor/master/preprocess.py --output knead.py
r="$?"
if [[ "$r" != "0" ]]; then
	echo "${red}Something went wrong with downloading the preprocessor${rst}"
	exit $r
fi
echo "--${green} Download latest preprocessor script ✅${rst}--"

echo "" 

echo "--${red} Install alias to .zshrc ${rst}--"
is_installed="$(cat ~/.zshrc | grep 'alias bakery')"
if [[ "$is_installed" == "" ]]; then
	echo 'alias bakery="python3 '"$HOME"'/Documents/bakery/knead.py"' >> "$HOME/.zshrc"
	source ~/.zshrc
else
	echo "Already installed"
fi
echo "--${green} Install alias to .zshrc ✅${rst}--"


echo ""
echo "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="

echo "${purp} All done! You can now prepare files for bakery 🍩 By doing the following: ${rst}"
echo ""
echo "- Drop the xlsx in this folder ($installDir)"
echo "- In a terminal, run ${purp}bakery${rst}"
echo ""
echo "You'll get a new file, dough.json ready to upload to bakery 🍩"
