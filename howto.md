https://login.microsoftonline.com/916fb24b-ab3d-4e9b-b76c-627574ca64a9/oauth2/v2.0/authorize?client_id=4f188305-3c5d-4e9d-8a94-11de8d1aa736&response_type=code&redirect_uri=http%3A%2F%2Flocalhost%3A6274%2Foauth%2Fcallback&response_mode=query&scope=offline_access%20User.Read%20Files.Read%20Mail.Read

--> copy code from code=<VALUE>

http://localhost:6274/oauth/callback?code= XXXXXXXX &session_state=008cde99-b115-5fb0-da7a-4354a0f81a4c#



>> DEFINITELY ADD SECRET AND CODE (RESULT CONTAINS ACCES AND REFRESH TOKEN >> we need refresh)
curl -X POST https://login.microsoftonline.com/916fb24b-ab3d-4e9b-b76c-627574ca64a9/oauth2/v2.0/token \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -d "grant_type=authorization_code" \
  -d "client_id=4f188305-3c5d-4e9d-8a94-11de8d1aa736" \
  -d "client_secret=YOUR_CLIENT_SECRET" \
  -d "code=PASTE_CODE_HERE" \
  -d "redirect_uri=http://localhost:6274/oauth/callback"



start the mcp:

npx tsx src/index.ts --http 3000 -v

connect openwebui to: 

http://host.docker.internal:3000/mcp


get a client id:
curl -sS -X POST http://localhost:3000/register \
  -H 'Content-Type: application/json' \
  -d '{"client_name":"OpenWebUI","redirect_uris":["http://localhost:8080/callback"],"grant_types":["authorization_code","refresh_token"],"response_types":["code"]}'

  