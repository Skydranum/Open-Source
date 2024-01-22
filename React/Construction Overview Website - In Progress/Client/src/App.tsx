import { Routes, Route } from "react-router-dom";
import { CSSProperties } from 'react';
import Footer from "./components/layouts/Footer";
import Navbar from "./components/layouts/Navbar";
import Home from "./components/pages/Home";
import Gfip from "./components/pages/Gfip";
import Mapa from "./components/pages/Mapa";
import ScrollToTop from './components/actions/Scroll';

function App() {
  const appStyle: CSSProperties = {
    background: 'linear-gradient(to right, #55852a, #5e8a37, #5e8a37, #5e8a37, #55852a)',
    minHeight: '50rem',
    overflow: 'auto',
    color: 'white',
    flex: '1 0 auto',
  };

  return (
    <>
      <div>
        <Navbar />
        <ScrollToTop />
        <div style={appStyle}>
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/gfip" element={<Gfip />} />
            <Route path="/mapa" element={<Mapa />} />
          </Routes>
        </div>
        <div style={{ flex: 1, flexShrink: 0 }}>
          <Footer />
        </div>
      </div>
    </>
  );
}

export default App