import { Container, Navbar } from 'react-bootstrap';
import ButtonTop from '../actions/ButtonTop';

function FooterMain() {
  return (
    <footer>
      <Navbar data-bs-theme="dark" expand="xl" className="bg-dark shadow-md">
        <Container className="d-flex justify-content-end">
          <Navbar.Text className="me-auto">© Copyright 2023 - Desenvolvido por Felipe Silveira. Todos direitos reservado</Navbar.Text>
          <ButtonTop />
        </Container>
      </Navbar>
    </footer>
  );
}

export default FooterMain;