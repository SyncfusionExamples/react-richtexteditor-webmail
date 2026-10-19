import { StoreProvider } from './store/Store';
import { MainLayout } from './shared/layouts/MainLayout';
import { AuthProvider } from './auth/AuthProvider';
import './App.css';

function App() {
  return (
    <StoreProvider>
      <AuthProvider>
        <MainLayout />
      </AuthProvider>
    </StoreProvider>
  );
}

export default App;


// import { useEffect } from "react";
// import { StoreProvider, useMailStore } from "./store/Store";
// import { mailAPI } from "./services/api";
// import AuthProvider from "./auth/AuthProvider";
// import { MainLayout } from "./shared/layouts/MainLayout";

// function AuthInitializer() {
//   const { loadRealData } = useMailStore();

//   useEffect(() => {
//     const params = new URLSearchParams(window.location.search);

//     const token = params.get("token");
//     const email = params.get("email");

//     if (token) {
//       console.log("✅ Token received from redirect");

//       // ✅ Store token
//       mailAPI.setToken(token);

//       // ✅ Store email (optional)
//       if (email) {
//         localStorage.setItem("authEmail", email);
//       }

//       // ✅ LOAD REAL DATA (THIS WAS MISSING ❗)
//       loadRealData(token);

//       // ✅ Clean URL
//       window.history.replaceState({}, document.title, "/");
//     }
//   }, [loadRealData]);

//   return null;
// }
// function App() {
//   return (
//     <StoreProvider>
//       <AuthProvider>
//         <AuthInitializer />  {/* ✅ ADD THIS LINE */}
//         <MainLayout />
//       </AuthProvider>
//     </StoreProvider>
//   );
// }
// export default App;
