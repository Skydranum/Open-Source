import { CSSProperties } from 'react';

export const Gfip = () => {
    const containerStyle: CSSProperties = {
        textAlign: 'center',
        padding: '20px', // Adjust as needed
    }

    const textStyle: CSSProperties = {
        maxWidth: '800px', // Adjust as needed
        margin: '0 auto', // Centers the text block
        marginBottom: '20px', // Add some space at the bottom
    }

    return (
        <>
            <div style={containerStyle}>
                <h3 style={textStyle}>GFIP</h3>
            </div>
        </>
    )
}

export default Gfip