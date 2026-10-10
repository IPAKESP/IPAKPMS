# IPAK PMS 2025 – Web System

Files
- `index.html`  – poora system (ek hi file). GitHub Pages / Netlify / kisi bhi hosting par rakh dein.
- `firestore.rules` – database security rules (Firebase mein paste karni hain).

## Local mode (bina setup ke)
`index.html` browser mein kholen. Data sirf usi browser mein save hota hai. Settings > "Local demo role" se har role try kar sakte hain.

## Cloud mode (Firebase, 10 minute, ek dafa)
1. console.firebase.google.com par project banayen > Web app add karein > `apiKey` aur `projectId` copy karein.
2. Authentication > Email/Password enable karein, har user ka email/password banayen.
3. Firestore Database banayen > Rules tab mein `firestore.rules` paste karke Publish karein.
4. Firestore mein collection `users` banayen. Har banday ka document ID = uski **lowercase email**, fields (string):
   - `role` = `admin` / `hr` / `lm` / `hod`
   - `person` = naam bilkul waisa jaisa employee list mein Supervisor/HOD column mein hai (sirf lm aur hod ke liye)
5. Authentication > Settings > Authorized domains mein apni hosting ka domain add karein.
6. `index.html` > Settings mein API key aur Project ID daalein > admin se sign in karein > "Upload all local data to cloud".

## Workflow
Line Manager form bhar ke submit > HOD approve/return > HR finalize (form lock) > HR employee acknowledgement record karta hai.
Har step audit trail mein save hota hai.

## Note
Role-based access UI mein poori tarah enforce hai. Firestore rules baseline protection deti hain (sirf logged-in users, role ke mutabiq write). Workflow ke har step ko server-side bhi lock karna ho to Cloud Functions chahiye hon ge.
