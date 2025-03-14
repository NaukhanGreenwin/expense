#!/bin/bash

# Create render.yaml
cat > render.yaml << 'EOL'
services:
  - type: web
    name: expense-report
    env: node
    buildCommand: npm install
    startCommand: node server.js
    envVars:
      - key: PORT
        value: 3005
      - key: OPENAI_API_KEY
        sync: false
EOL

# Add, commit, and push
git add render.yaml
git commit -m "Add Render configuration"
git push

echo "Deployment files pushed to GitHub."
echo "Now go to Render.com to deploy your application." 