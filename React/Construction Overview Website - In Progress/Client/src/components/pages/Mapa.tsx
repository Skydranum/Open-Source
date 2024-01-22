import React, { useRef, useState, useEffect } from 'react';
import { Button, Modal, Form } from 'react-bootstrap';
import mapboxgl from 'mapbox-gl';
import axios from 'axios';
import MapboxGeocoder from '@mapbox/mapbox-gl-geocoder';
import '@mapbox/mapbox-gl-geocoder/dist/mapbox-gl-geocoder.css';

type Point = {
    CoordenadasX: number;
    CoordenadasY: number;
    Nome: string;
    Descricao: string;
    id: number;
    Nomeclatura: string | null;
    PDF: string | null;
    Tag1: string | null;
    Tag2: string | null;
    Tag3: string | null;
    Tag4: string | null;
};

mapboxgl.accessToken = 'MAPBOX API KEY';

const Mapa = () => {
    const mapContainer = useRef<HTMLDivElement>(null);
    const map = useRef<mapboxgl.Map | null>(null);
    const [dataPoints, setDataPoints] = useState<Point[]>([]);
    const [isAddMode, setIsAddMode] = useState<boolean>(false);
    const [isDragMode, setIsDragMode] = useState<boolean>(false);
    const [showModal, setShowModal] = useState<boolean>(false);
    const [currentCoords, setCurrentCoords] = useState<[number, number] | null>(null);
    const [currentNome, setCurrentNome] = useState<string>('');
    const [currentDescricao, setCurrentDescricao] = useState<string>('');
    const [currentNomenclature, setCurrentNomenclature] = useState<string | null>(null);
    const [showEditModal, setShowEditModal] = useState<boolean>(false);
    const [editingPoint, setEditingPoint] = useState<Point | null>(null);
    const [markers, setMarkers] = useState<mapboxgl.Marker[]>([]);
    const [selectedFile, setSelectedFile] = useState<File | null>(null);
    const [currentTag1, setCurrentTag1] = useState<string>('');
    const [currentTag2, setCurrentTag2] = useState<string>('');
    const [currentTag3, setCurrentTag3] = useState<string>('');
    const [currentTag4, setCurrentTag4] = useState<string>('');
    const availableTags = ["Hélice Continua", "Metálica", "Sondagem", "Pré-Moldada", "Furo Teste", "Parede Diafragma", "Grampo", "Raiz", "Escavada", "Tirante"];
    const [currentAvailableTags, setCurrentAvailableTags] = useState(availableTags);
    const [filterTag, setFilterTag] = useState<string | null>(null);

    const createMarkers = (draggable: boolean) => {
        markers.forEach(marker => marker.remove());
        markers.length = 0;

        let pointsToDisplay = dataPoints;

        if (filterTag) {
            pointsToDisplay = pointsToDisplay.filter(point =>
                [point.Tag1, point.Tag2, point.Tag3, point.Tag4].includes(filterTag)
            );
        }

        pointsToDisplay.forEach(point => {
            const color = point.Nomeclatura === "Sondagem" ? "blue" :
                point.Nomeclatura === "Obra" ? "green" :
                    point.Nomeclatura === "Orçamento" ? "white" : "gray";
            const marker = new mapboxgl.Marker({ color: color, draggable: draggable })
                .setLngLat([point.CoordenadasX, point.CoordenadasY])
                .addTo(map.current!);

            if (draggable) {
                marker.on('dragend', (event) => {
                    const { lng, lat } = event.target.getLngLat();
                    axios.put(`http://localhost:3001/sondagens/${point.id}`, {
                        CoordenadasX: lng,
                        CoordenadasY: lat,
                        Nome: point.Nome,
                        Descricao: point.Descricao,
                        Nomeclatura: point.Nomeclatura,
                        PDF: point.PDF
                    }).then(() => {
                        setIsDragMode(false);  // Desativa o modo de arrasto
                        fetchDataAndUpdateMap(); // Recarrega os dados e atualiza os pins
                    });
                });
            }

            markers.push(marker);
            const pdfUrl = `http://localhost:3001${point.PDF}`;
            const popup = new mapboxgl.Popup({ offset: 25 })
                .setHTML(`
                <span style="color: black;">
                    Nome: ${point.Nome}<br>
                    Descrição: ${point.Descricao}<br>
                    Serviço: ${point.Nomeclatura}<br>
                    Tags: ${point.Tag1 ? point.Tag1 + ' - ' : ''}${point.Tag2 ? point.Tag2 + ' - ' : ''}${point.Tag3 ? point.Tag3 + ' - ' : ''}${point.Tag4 || ''}<br>
                    <button data-id="${point.id}" class="edit-pin-btn">Editar</button>
                    <button data-id="${point.id}" class="delete-pin-btn">Deletar</button>
                    <a href="${pdfUrl}" target="_blank">Ver PDF</a>
                </span>
            `);
            marker.setPopup(popup);
        });
        setMarkers(markers);
    };

    const fetchDataAndUpdateMap = () => {
        axios.get('http://localhost:3001/sondagens').then(response => {
            const newDataPoints = response.data as Point[];
            setDataPoints(newDataPoints);
            createMarkers(isDragMode);
        });
    };

    useEffect(() => {
        createMarkers(isDragMode);
    }, [isDragMode, dataPoints, filterTag]); // Adiciona Tag aqui

    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (event.target.files && event.target.files.length > 0) {
            setSelectedFile(event.target.files[0]);
        }
    };

    const handleMapClick = (e: mapboxgl.MapMouseEvent) => {
        if (isAddMode) {
            setCurrentCoords([e.lngLat.lng, e.lngLat.lat]);
            setShowModal(true);
        }
    };
    useEffect(() => {
        document.addEventListener('click', handleDocumentClick);

        return () => {
            document.removeEventListener('click', handleDocumentClick);
        };
    }, [dataPoints]);

    useEffect(() => {
        map.current = new mapboxgl.Map({
            container: mapContainer.current!,
            style: 'mapbox://styles/mapbox/satellite-v9',
            center: [0, 0],
            zoom: 2,
        });

        // Adicione o Geocoder ao mapa
        const geocoder = new MapboxGeocoder({
            accessToken: mapboxgl.accessToken,
            mapboxgl: mapboxgl,
            placeholder: 'Procure um local...',
        });

        map.current.addControl(geocoder);

        // Move o mapa para o local selecionado
        geocoder.on('result', (e) => {
            map.current!.flyTo({ center: e.result.geometry.coordinates });
        });

        fetchDataAndUpdateMap();

        return () => {
            map.current?.remove();
        };
    }, []);

    useEffect(() => {
        if (map.current) {
            if (isAddMode) {
                map.current.on('click', handleMapClick);
            } else {
                map.current.off('click', handleMapClick);
            }
        }

        return () => {
            if (map.current) {
                map.current.off('click', handleMapClick);
            }
        };
    }, [isAddMode]);

    const handleSave = () => {
        if (currentCoords) {
            const [CoordenadasX, CoordenadasY] = currentCoords;

            if (selectedFile) {
                const formData = new FormData();
                console.log("Sending file to server:", selectedFile.name);
                formData.append('pdfFile', selectedFile);

                axios.post('http://localhost:3001/upload', formData).then(response => {
                    const filePath = response.data.filePath;
                    savePoint(CoordenadasX, CoordenadasY, filePath);
                });
            } else {
                savePoint(CoordenadasX, CoordenadasY, null);
            }
        }
    };

    const savePoint = (CoordenadasX: number, CoordenadasY: number, filePath: string | null) => {
        axios.post('http://localhost:3001/sondagens', {
            CoordenadasX, CoordenadasY, Nome: currentNome, Descricao: currentDescricao, Nomeclatura: currentNomenclature, filePath, Tag1: currentTag1, Tag2: currentTag2, Tag3: currentTag3, Tag4: currentTag4
        }).then(() => {
            setShowModal(false);
            fetchDataAndUpdateMap();
            setIsAddMode(false);
        });
    };

    const resetModalData = () => {
        setCurrentNome('');
        setCurrentDescricao('');
        setCurrentNomenclature(null);
        setSelectedFile(null);
        setCurrentTag1('');
        setCurrentTag2('');
        setCurrentTag3('');
        setCurrentTag4('');
        setCurrentAvailableTags(availableTags);
    };

    const initiateDelete = (id: number) => {
        axios.delete(`http://localhost:3001/sondagens/${id}`).then(() => {
            setDataPoints(prevDataPoints => prevDataPoints.filter(point => point.id !== id));
            fetchDataAndUpdateMap();
        });
    };

    const handleDocumentClick = (e: Event) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains("edit-pin-btn")) {
            const id = parseInt(target.getAttribute("data-id")!, 10);
            openEditModal(id);
        } else if (target.classList.contains("delete-pin-btn")) {
            const id = parseInt(target.getAttribute("data-id")!, 10);
            initiateDelete(id);
        }
    };

    const handleEditSave = () => {
        if (editingPoint && currentNome && currentDescricao) {
            const formData = new FormData();
            formData.append('Nome', currentNome);
            formData.append('Descricao', currentDescricao);
            formData.append('Nomeclatura', currentNomenclature || '');
            formData.append('Tag1', currentTag1);
            formData.append('Tag2', currentTag2);
            formData.append('Tag3', currentTag3);
            formData.append('Tag4', currentTag4);

            if (selectedFile) {
                formData.append('pdfFile', selectedFile);
            }

            axios.put(`http://localhost:3001/sondagens/${editingPoint.id}`, formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            }).then(() => {
                setShowEditModal(false);
                fetchDataAndUpdateMap();
                setSelectedFile(null);
            });
        }
    };

    const openEditModal = (id: number) => {
        const pointToEdit = dataPoints.find(point => point.id === id);
        if (pointToEdit) {
            setCurrentNome(pointToEdit.Nome);
            setCurrentDescricao(pointToEdit.Descricao);
            setCurrentNomenclature(pointToEdit.Nomeclatura);
            setCurrentTag1(pointToEdit.Tag1 || '');
            setCurrentTag2(pointToEdit.Tag2 || '');
            setCurrentTag3(pointToEdit.Tag3 || '');
            setCurrentTag4(pointToEdit.Tag4 || '');
            setEditingPoint(pointToEdit);
            setShowEditModal(true);
        }
        const currentTags = [pointToEdit.Tag1, pointToEdit.Tag2, pointToEdit.Tag3, pointToEdit.Tag4].filter(Boolean) as string[]; // Não sei como isso está funcionando.
        const available = availableTags.filter(tag => !currentTags.includes(tag));
        setCurrentAvailableTags(available);
    };

    return (
        <>
            <Button onClick={() => {
                resetModalData();
                if (isAddMode) {
                    setIsAddMode(false);
                    setShowModal(false); // Isso garante que o modal seja fechado se isAddMode for true
                } else {
                    setIsAddMode(true);
                }
            }} style={{ marginLeft: '10px' }}>
                {isAddMode ? "Cancelar" : "Criar um novo ponto"}
            </Button>

            <Button onClick={() => setIsDragMode(!isDragMode)} style={{ marginLeft: '10px' }}>
                {isDragMode ? "Desativar" : "Editar Localização"}
            </Button>

            <label style={{ marginLeft: '10px' }} >Filtrar por Tag:  </label>
            <select value={filterTag || ''} onChange={(e) => setFilterTag(e.target.value || null)}>
                <option value="">Mostrar Todos</option>
                {availableTags.map(tag => (
                    <option key={tag} value={tag}>{tag}</option>
                ))}
            </select>

            <Modal show={showEditModal || showModal} onHide={() => { setShowModal(false); setShowEditModal(false); }}>
                <Modal.Header closeButton>
                    <Modal.Title>{showEditModal ? "Editar Ponto" : "Adicionar Ponto"}</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    <Form>
                        <Form.Group>
                            <Form.Label>Nome</Form.Label>
                            <Form.Control type="text" value={currentNome} onChange={(e) => setCurrentNome(e.target.value)} />
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Descrição</Form.Label>
                            <Form.Control type="text" value={currentDescricao} onChange={(e) => setCurrentDescricao(e.target.value)} />
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Nomenclatura</Form.Label>
                            <Form.Control as="select" value={currentNomenclature || ''} onChange={(e) => setCurrentNomenclature(e.target.value)}>
                                <option value="">Nenhum</option>
                                <option value="Sondagem">Sondagem</option>
                                <option value="Obra">Obra</option>
                                <option value="Orçamento">Orçamento</option>
                            </Form.Control>
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Upload PDF</Form.Label>
                            <Form.Control type="file" onChange={handleFileChange} />
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Tag 1</Form.Label>
                            <Form.Control as="select" value={currentTag1} onChange={(e) => {
                                setCurrentTag1(e.target.value);
                                const otherSelectedTags = [e.target.value, currentTag2, currentTag3, currentTag4];
                                setCurrentAvailableTags(prevTags => availableTags.filter(tag => !otherSelectedTags.includes(tag)));
                            }}>

                                <option value="">Selecione uma tag</option>
                                {currentAvailableTags.concat(currentTag1).sort().filter((value, index, self) => self.indexOf(value) === index).map(tag =>
                                    <option key={tag} value={tag}>{tag}</option>
                                )}
                            </Form.Control>
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Tag 2</Form.Label>
                            <Form.Control as="select" value={currentTag2} onChange={(e) => {
                                setCurrentTag2(e.target.value);
                                const otherSelectedTags = [e.target.value, currentTag1, currentTag3, currentTag4];
                                setCurrentAvailableTags(prevTags => availableTags.filter(tag => !otherSelectedTags.includes(tag)));
                            }}>

                                <option value="">Selecione uma tag</option>
                                {currentAvailableTags.concat(currentTag2).sort().filter((value, index, self) => self.indexOf(value) === index).map(tag =>
                                    <option key={tag} value={tag}>{tag}</option>
                                )}
                            </Form.Control>
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Tag 3</Form.Label>
                            <Form.Control as="select" value={currentTag3} onChange={(e) => {
                                setCurrentTag1(e.target.value);
                                const otherSelectedTags = [e.target.value, currentTag1, currentTag2, currentTag4];
                                setCurrentAvailableTags(prevTags => availableTags.filter(tag => !otherSelectedTags.includes(tag)));
                            }}>

                                <option value="">Selecione uma tag</option>
                                {currentAvailableTags.concat(currentTag3).sort().filter((value, index, self) => self.indexOf(value) === index).map(tag =>
                                    <option key={tag} value={tag}>{tag}</option>
                                )}
                            </Form.Control>
                        </Form.Group>
                        <Form.Group>
                            <Form.Label>Tag 4</Form.Label>
                            <Form.Control as="select" value={currentTag4} onChange={(e) => {
                                setCurrentTag1(e.target.value);
                                const otherSelectedTags = [e.target.value, currentTag1, currentTag2, currentTag3];
                                setCurrentAvailableTags(prevTags => availableTags.filter(tag => !otherSelectedTags.includes(tag)));
                            }}>

                                <option value="">Selecione uma tag</option>
                                {currentAvailableTags.concat(currentTag4).sort().filter((value, index, self) => self.indexOf(value) === index).map(tag =>
                                    <option key={tag} value={tag}>{tag}</option>
                                )}
                            </Form.Control>
                        </Form.Group>
                    </Form>
                </Modal.Body>
                <Modal.Footer>
                    <Button variant="secondary" onClick={() => { setShowModal(false); setShowEditModal(false); }}>Fechar</Button>
                    <Button variant="primary" onClick={showEditModal ? handleEditSave : handleSave}>
                        {showEditModal ? "Salvar Mudanças" : "Salvar Mudanças"}
                    </Button>
                </Modal.Footer>
            </Modal>

            <div ref={mapContainer} style={{ width: '100vw', height: '100vh' }} />
        </>
    );
};

export default Mapa;