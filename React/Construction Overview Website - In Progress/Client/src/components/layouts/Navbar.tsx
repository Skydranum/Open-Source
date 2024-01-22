import { Navbar as NavbarBS, Nav, NavDropdown } from 'react-bootstrap';
import { Link } from "react-router-dom";
import { Cash, EmojiSmile, Gear } from 'react-bootstrap-icons'

function NavbarMain() {
  return (
    <>
      <NavbarBS sticky="top" data-bs-theme="dark" expand="lg" className="bg-dark shadow-md me-auto">
        <NavbarBS.Brand as={Link} to="/">
          <img
            src="./icons/logo.png"
            width="35"
            height="35"
            className="d-inline-block align-top ms-3"
            style={{ backgroundColor: 'white' }}
          />
        </NavbarBS.Brand>
        <NavbarBS.Toggle aria-controls="basic-navbar-nav" className="me-2" />
        <NavbarBS.Collapse id="basic-navbar-nav" className="justify-content-between ms-1">
          <Nav>

            <NavDropdown title={<span><Cash className="mb-1 me-1" />Comercial</span>}>
              <NavDropdown.Item as={Link} to="/propostas">
                <span className="d-flex align-items-center">Propostas</span>
              </NavDropdown.Item>
            </NavDropdown>

            <NavDropdown title={<span><Gear className="mb-1 me-1" />Produção</span>}>
              <NavDropdown.Item as={Link} to="/obras">
                <span className="d-flex align-items-center">Obras</span>
              </NavDropdown.Item>
              <NavDropdown.Item as={Link} to="/mapa">
                <span className="d-flex align-items-center">Mapa</span>
              </NavDropdown.Item>
            </NavDropdown>

            <NavDropdown title={<span><EmojiSmile className="mb-1 me-1" />Funcionários</span>}>
              <NavDropdown.Item as={Link} to="/gfip">
                <span className="d-flex align-items-center">Gfip</span>
              </NavDropdown.Item>
            </NavDropdown>

          </Nav>
        </NavbarBS.Collapse>
      </NavbarBS>
    </>
  );
}

export default NavbarMain