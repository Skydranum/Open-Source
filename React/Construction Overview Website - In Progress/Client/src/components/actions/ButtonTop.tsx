import { Button } from 'react-bootstrap';
import { ArrowUp } from 'react-bootstrap-icons';

function ButtonTop() {
  const scrollToTop = () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth"  // for smooth scrolling
    });
  };

  return (
    <Button onClick={scrollToTop}>
      <ArrowUp className='me-2' />Voltar para cima
    </Button>
  );
}

export default ButtonTop;
