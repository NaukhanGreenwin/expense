#!/bin/bash

# Add the modified files
git add public/js/app.js public/css/styles.css

# Commit the changes
git commit -m "Add custom G/L code input when Other is selected"

# Push to GitHub
git push

echo "Changes pushed to GitHub. Render.com will automatically deploy the updates." 