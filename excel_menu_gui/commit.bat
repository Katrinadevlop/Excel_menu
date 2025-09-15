@echo off
echo Adding all files to git...
git add -A

echo.
echo Creating commit...
git commit -m "Add extract_fish_dishes_by_range function and update extract_fish_dishes_from_excel to use column E, F, G like other categories"

echo.
echo Commit complete!
pause
