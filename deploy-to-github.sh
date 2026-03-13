#!/bin/bash
# Deploy Y2K Portfolio to GitHub

set -e

cd /git/y2k-portfolio

# Extract PAT from git-credentials
TOKEN=$(grep github /root/.git-credentials | sed 's/.*:\/\/[^:]*:\([^@]*\)@.*/\1/')

echo "Creating GitHub repository..."
curl -s -H "Authorization: token $TOKEN" \
  https://api.github.com/user/repos \
  -d '{
    "name":"y2k-portfolio",
    "description":"Y2K Global & KISS IT Solutions - AI Portfolio & Marketing Materials",
    "private":false
  }' > repo_creation.log

echo "Pushing to GitHub..."
git push -u origin main

echo "Inviting gerykiss86 as admin collaborator..."
curl -s -H "Authorization: token $TOKEN" \
  https://api.github.com/repos/worker1-kiss-it/y2k-portfolio/collaborators/gerykiss86 \
  -X PUT -d '{"permission":"admin"}' > invite_collaborator.log

echo "Deployment complete!"
echo "Repository URL: https://github.com/worker1-kiss-it/y2k-portfolio"
echo "Website URL: https://worker1-kiss-it.github.io/y2k-portfolio (if GitHub Pages enabled)"
