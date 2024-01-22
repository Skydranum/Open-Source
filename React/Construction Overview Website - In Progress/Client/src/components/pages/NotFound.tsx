import { CSSProperties } from 'react';

export const About = () => {
    const containerStyle: CSSProperties = {
        textAlign: 'center',
        padding: '20px',
    }

    const textStyle: CSSProperties = {
        maxWidth: '800px',
        margin: '0 auto',
        marginBottom: '20px',
    }

    return (
        <div style={containerStyle}>
            <h3 style={textStyle}>Pagina não encontrada</h3>
        </div>
    )
}

export default About