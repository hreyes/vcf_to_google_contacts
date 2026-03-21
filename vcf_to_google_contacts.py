#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para convertir archivos .vcf (vCard) a CSV compatible con Google Contacts.
Maneja duplicados, normaliza nombres y unifica múltiples contactos.
"""

import csv
import os
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Set


class VCardParser:
    """Parser para archivos vCard (.vcf).
    
    Procesa archivos en formato vCard (RFC 6350) y extrae información de contactos.
    Soporta encoding QUOTED-PRINTABLE y líneas continuadas.
    """

    def __init__(self, vcf_file: str) -> None:
        """Inicializa el parser.
        
        Args:
            vcf_file: Ruta al archivo vCard (.vcf)
        """
        self.vcf_file = vcf_file
        self.contacts: List[Dict] = []

    def parse(self) -> List[Dict]:
        """Lee y parsea el archivo .vcf completo.
        
        Returns:
            Lista de diccionarios con información de contactos.
            
        Raises:
            FileNotFoundError: Si el archivo no existe.
            PermissionError: Si no hay permisos de lectura.
            ValueError: Si el archivo está vacío.
        """
        # Validación del archivo
        vcf_path = Path(self.vcf_file)
        
        if not vcf_path.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {self.vcf_file}")
        
        if not vcf_path.is_file():
            raise ValueError(f"La ruta no es un archivo: {self.vcf_file}")
        
        if not os.access(self.vcf_file, os.R_OK):
            raise PermissionError(f"Permiso denegado al leer: {self.vcf_file}")
        
        if vcf_path.stat().st_size == 0:
            raise ValueError(f"El archivo está vacío: {self.vcf_file}")
        
        try:
            with open(self.vcf_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except Exception as e:
            raise RuntimeError(f"Error crítico al leer el archivo: {e}")

        # Unir líneas continuadas que empiezan con espacio o tab en todo el archivo
        content = re.sub(r'\r?\n[ \t]', '', content)

        # Dividir en vCards individuales
        vcards = content.split('END:VCARD')

        for vcard_text in vcards:
            if 'BEGIN:VCARD' not in vcard_text:
                continue

            contact = self._parse_vcard(vcard_text)
            if contact:
                self.contacts.append(contact)

        return self.contacts

    def _decode_value(self, value: str, params: List[str]) -> str:
        """Decodifica un valor si está marcado como QUOTED-PRINTABLE.
        
        Args:
            value: Valor a decodificar.
            params: Parámetros del campo vCard.
            
        Returns:
            Valor decodificado.
        """
        is_quoted = any('ENCODING=QUOTED-PRINTABLE' in p.upper() for p in params)

        if is_quoted:
            try:
                import quopri
                return quopri.decodestring(value).decode('utf-8', 'ignore')
            except Exception:
                return value

        return value.replace('\\n', '\n').replace('\\,', ',').replace('\\;', ';').strip()

    def _parse_vcard(self, vcard_text: str) -> Optional[Dict]:
        """Parsea un vCard individual y extrae todos los campos.
        
        Args:
            vcard_text: Texto del vCard a parsear.
            
        Returns:
            Diccionario con datos del contacto o None si está vacío.
        """
        contact = {
            'fn': '', 'family_name': '', 'given_name': '', 'middle_name': '',
            'phones': [], 'emails': [], 'notes': '', 'org': '', 'addresses': [], 'photo': ''
        }

        lines = vcard_text.strip().split('\n')

        for line in lines:
            line = line.strip()
            if not line or ':' not in line:
                continue

            field_part, _, value = line.partition(':')
            field_parts = field_part.split(';')
            field = field_parts[0].upper()
            params = field_parts[1:]

            decoded_value = self._decode_value(value, params)

            if field == 'FN':
                contact['fn'] = decoded_value
            elif field == 'N':
                parts = decoded_value.split(';')
                contact['family_name'] = parts[0] if len(parts) > 0 else ''
                contact['given_name'] = parts[1] if len(parts) > 1 else ''
                contact['middle_name'] = parts[2] if len(parts) > 2 else ''
            elif field == 'TEL':
                phone = self._clean_phone(decoded_value)
                if phone:
                    contact['phones'].append({'number': phone, 'type': self._extract_type(params)})
            elif field == 'EMAIL':
                if '@' in decoded_value:
                    contact['emails'].append({'address': decoded_value, 'type': self._extract_type(params)})
            elif field == 'NOTE':
                contact['notes'] = (contact['notes'] + ' | ' + decoded_value) if contact['notes'] else decoded_value
            elif field == 'ORG':
                contact['org'] = decoded_value
            elif field == 'ADR':
                addr = decoded_value.replace(';', ', ')
                if addr:
                    contact['addresses'].append({'address': addr, 'type': self._extract_type(params)})
            elif field == 'PHOTO':
                contact['photo'] = value[:100]

        contact['fn'] = self._normalize_full_name(contact)
        if contact['fn'] or contact['phones']:
            return contact
        return None

    def _extract_type(self, params: List[str]) -> str:
        """Extrae el tipo de un campo (CELL, HOME, WORK, etc.).
        
        Args:
            params: Parámetros del campo vCard.
            
        Returns:
            Tipo del campo capitalizado (ej: 'Cell', 'Home', 'Work').
        """
        for param in params:
            if param.upper().startswith('TYPE='):
                return param.split('=')[1].capitalize()
            elif param.upper() in ['CELL', 'HOME', 'WORK', 'VOICE', 'FAX', 'PREF', 'MAIN']:
                return param.capitalize()
        return 'Other'

    def _clean_phone(self, phone: str) -> Optional[str]:
        """Limpia y normaliza un número de teléfono.
        
        Args:
            phone: Número de teléfono a limpiar.
            
        Returns:
            Teléfono limpio o None si está vacío.
        """
        phone = re.sub(r'[^\d+]', '', phone)
        return phone if phone else None

    def _normalize_full_name(self, contact: Dict) -> str:
        """Normaliza el nombre completo del contacto.
        
        Args:
            contact: Diccionario con datos del contacto.
            
        Returns:
            Nombre completo normalizado.
        """
        if contact['fn']:
            return contact['fn']
        parts = []
        if contact['given_name']: parts.append(contact['given_name'])
        if contact['middle_name']: parts.append(contact['middle_name'])
        if contact['family_name']: parts.append(contact['family_name'])
        if parts:
            return ' '.join(parts)
        if contact['phones']:
            return f"Contacto {contact['phones'][0]['number']}"
        return "Sin nombre"


class ContactMerger:
    """Unifica contactos duplicados basándose en nombre o teléfono.
    
    Detecta y fusiona contactos duplicados usando nombre y teléfono como criterios.
    Mantiene la información más completa de cada grupo de duplicados.
    """

    def __init__(self, contacts: List[Dict]) -> None:
        """Inicializa el unificador.
        
        Args:
            contacts: Lista de contactos a procesar.
        """
        self.contacts = contacts

    def merge_duplicates(self) -> List[Dict]:
        """Fusiona contactos duplicados.
        
        Detecta duplicados usando nombre y números de teléfono como criterios.
        
        Returns:
            Lista de contactos sin duplicados.
        """
        # Índices por teléfono y por nombre
        phone_index: Dict[str, List[int]] = defaultdict(list)
        name_index: Dict[str, List[int]] = defaultdict(list)

        # Construir índices
        for idx, contact in enumerate(self.contacts):
            # Indexar por nombre
            name_key = contact['fn'].lower().strip()
            if name_key and name_key != 'sin nombre':
                name_index[name_key].append(idx)

            # Indexar por teléfonos
            for phone in contact['phones']:
                phone_index[phone['number']].append(idx)

        # Encontrar grupos de contactos a fusionar
        merged_groups: List[Set[int]] = []
        processed = set()

        for indices in phone_index.values():
            if len(indices) > 1:
                # Múltiples contactos con el mismo teléfono
                group = set(indices)
                merged_groups.append(group)
                processed.update(group)

        for indices in name_index.values():
            if len(indices) > 1:
                # Múltiples contactos con el mismo nombre
                group = set(indices)
                # Ver si ya están en un grupo
                merged = False
                for existing_group in merged_groups:
                    if group & existing_group:
                        existing_group.update(group)
                        merged = True
                        break
                if not merged:
                    merged_groups.append(group)
                processed.update(group)

        # Fusionar grupos
        merged_contacts = []

        for group in merged_groups:
            merged_contact = self._merge_group([self.contacts[i] for i in group])
            merged_contacts.append(merged_contact)

        # Agregar contactos no fusionados
        for idx, contact in enumerate(self.contacts):
            if idx not in processed:
                merged_contacts.append(contact)

        return merged_contacts

    def _merge_group(self, group: List[Dict]) -> Dict:
        """Fusiona un grupo de contactos en uno solo.
        
        Combina la información de múltiples contactos, eligiendo los datos
        más completos y eliminando duplicados de teléfono y email.
        
        Args:
            group: Lista de contactos a fusionar.
            
        Returns:
            Contacto único con información combinada.
        """
        merged = {
            'fn': '',
            'family_name': '',
            'given_name': '',
            'middle_name': '',
            'phones': [],
            'emails': [],
            'notes': '',
            'org': '',
            'addresses': [],
            'photo': ''
        }

        seen_phones = set()
        seen_emails = set()
        notes_parts = []
        orgs = []

        for contact in group:
            # Nombre: usar el más completo
            if len(contact['fn']) > len(merged['fn']):
                merged['fn'] = contact['fn']
                merged['given_name'] = contact['given_name']
                merged['family_name'] = contact['family_name']
                merged['middle_name'] = contact['middle_name']

            # Teléfonos: agregar únicos
            for phone in contact['phones']:
                if phone['number'] not in seen_phones:
                    merged['phones'].append(phone)
                    seen_phones.add(phone['number'])

            # Emails: agregar únicos
            for email in contact['emails']:
                if email['address'].lower() not in seen_emails:
                    merged['emails'].append(email)
                    seen_emails.add(email['address'].lower())

            # Notas: concatenar
            if contact['notes']:
                notes_parts.append(contact['notes'])

            # Organización
            if contact['org'] and contact['org'] not in orgs:
                orgs.append(contact['org'])

            # Direcciones
            for addr in contact['addresses']:
                if addr not in merged['addresses']:
                    merged['addresses'].append(addr)

            # Foto
            if contact['photo'] and not merged['photo']:
                merged['photo'] = contact['photo']

        # Unir notas
        merged['notes'] = ' | '.join(notes_parts)

        # Unir organizaciones
        merged['org'] = ', '.join(orgs)

        return merged


class GoogleContactsCSV:
    """Genera CSV compatible con Google Contacts.
    
    Convierte una lista de contactos a formato CSV que puede importarse
    directamente en Google Contacts.
    """

    def __init__(self, contacts: List[Dict], output_file: str) -> None:
        """Inicializa el generador CSV.
        
        Args:
            contacts: Lista de contactos a exportar.
            output_file: Ruta del archivo CSV de salida.
        """
        self.contacts = contacts
        self.output_file = output_file
        # Encabezados EXACTOS proporcionados por el usuario
        self.headers = [
            'First Name', 'Middle Name', 'Last Name', 'Phonetic First Name',
            'Phonetic Middle Name', 'Phonetic Last Name', 'Name Prefix', 'Name Suffix',
            'Nickname', 'File As', 'Organization Name', 'Organization Title',
            'Organization Department', 'Birthday', 'Notes', 'Photo', 'Labels',
            'E-mail 1 - Label', 'E-mail 1 - Value', 'Phone 1 - Label', 'Phone 1 - Value',
            'Phone 2 - Label', 'Phone 2 - Value', 'Phone 3 - Label', 'Phone 3 - Value',
            'Address 1 - Label', 'Address 1 - Formatted', 'Address 1 - Street',
            'Address 1 - City', 'Address 1 - PO Box', 'Address 1 - Region',
            'Address 1 - Postal Code', 'Address 1 - Country', 'Address 1 - Extended Address',
            'Custom Field 1 - Label', 'Custom Field 1 - Value'
        ]

    def generate(self) -> None:
        """Genera el archivo CSV.
        
        Escribe los contactos en formato CSV compatible con Google Contacts.
        
        Raises:
            IOError: Si no se puede escribir el archivo de salida.
            ValueError: Si no hay contactos para exportar.
        """
        if not self.contacts:
            raise ValueError("No hay contactos para exportar")
        
        try:
            # Validar que el directorio de salida existe
            output_path = Path(self.output_file)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=self.headers)
                writer.writeheader()

                for contact in self.contacts:
                # Mapeo a las nuevas columnas
                row = {
                    'First Name': contact.get('given_name', ''),
                    'Middle Name': contact.get('middle_name', ''),
                    'Last Name': contact.get('family_name', ''),
                    'Organization Name': contact.get('org', ''),
                    'Notes': contact.get('notes', ''),
                    'Photo': contact.get('photo', '')
                }

                # Mapear hasta 3 teléfonos
                for i, phone in enumerate(contact.get('phones', [])[:3], 1):
                    row[f'Phone {i} - Label'] = phone.get('type', 'Other')
                    row[f'Phone {i} - Value'] = phone.get('number', '')

                # Mapear hasta 1 email (se puede extender si es necesario)
                for i, email in enumerate(contact.get('emails', [])[:1], 1):
                    row[f'E-mail {i} - Label'] = email.get('type', 'Other')
                    row[f'E-mail {i} - Value'] = email.get('address', '')

                # Mapear primera dirección
                if contact.get('addresses'):
                    address = contact['addresses'][0]
                    row['Address 1 - Label'] = address.get('type', 'Home')
                    row['Address 1 - Formatted'] = address.get('address', '')

                    writer.writerow(row)
        except IOError as e:
            raise IOError(f"Error al escribir el archivo CSV: {e}")
        except Exception as e:
            raise RuntimeError(f"Error inesperado al generar CSV: {e}")

        print(f"CSV generado exitosamente: {self.output_file}")


def main() -> None:
    """Función principal.
    
    Orquesta el flujo de conversión: parseo, unificación y generación de CSV.
    """
    print("=" * 70)
    print("Conversor de vCard (.vcf) a Google Contacts CSV")
    print("=" * 70)

    # Archivo de entrada
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = 'vcards_20260115_140358.vcf'

    # Archivo de salida
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    else:
        output_file = 'salida.csv'

    print(f"\nArchivo de entrada: {input_file}")
    print(f"Archivo de salida: {output_file}\n")

    # 1. Parsear el archivo VCF
    print("Paso 1/3: Parseando archivo .vcf...")
    try:
        parser = VCardParser(input_file)
        contacts = parser.parse()
        print(f"[OK] {len(contacts)} contactos extraídos")
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except PermissionError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Error al procesar archivo: {e}")
        sys.exit(1)

    # 2. Fusionar duplicados
    print("\nPaso 2/3: Fusionando contactos duplicados...")
    merger = ContactMerger(contacts)
    merged_contacts = merger.merge_duplicates()
    duplicates_removed = len(contacts) - len(merged_contacts)
    print(f"[OK] {duplicates_removed} duplicados fusionados")
    print(f"[OK] {len(merged_contacts)} contactos únicos")

    # 3. Generar CSV
    print("\nPaso 3/3: Generando CSV compatible con Google Contacts...")
    try:
        csv_generator = GoogleContactsCSV(merged_contacts, output_file)
        csv_generator.generate()
    except ValueError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except IOError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Error al generar CSV: {e}")
        sys.exit(1)

    # Estadísticas finales
    print("\n" + "=" * 70)
    print("ESTADISTICAS:")
    print(f"   - Contactos procesados: {len(contacts)}")
    print(f"   - Duplicados fusionados: {duplicates_removed}")
    print(f"   - Contactos en CSV final: {len(merged_contacts)}")

    total_phones = sum(len(c['phones']) for c in merged_contacts)
    total_emails = sum(len(c['emails']) for c in merged_contacts)
    print(f"   - Total telefonos: {total_phones}")
    print(f"   - Total emails: {total_emails}")

    print("\nProceso completado exitosamente!")
    print(f"Puedes importar {output_file} directamente en Google Contacts")
    print("=" * 70)


if __name__ == '__main__':
    main()
