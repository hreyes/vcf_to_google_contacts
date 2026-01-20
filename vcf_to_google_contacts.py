#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para convertir archivos .vcf (vCard) a CSV compatible con Google Contacts.
Maneja duplicados, normaliza nombres y unifica múltiples contactos.
"""

import csv
import re
import sys
from collections import defaultdict
from typing import Dict, List, Set


class VCardParser:
    """Parser para archivos vCard (.vcf)"""

    def __init__(self, vcf_file: str):
        self.vcf_file = vcf_file
        self.contacts = []

    def parse(self) -> List[Dict]:
        """Lee y parsea el archivo .vcf completo"""
        try:
            with open(self.vcf_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except UnicodeDecodeError:
            # Intentar con otras codificaciones comunes
            with open(self.vcf_file, 'r', encoding='latin-1', errors='ignore') as f:
                content = f.read()

        # Dividir en vCards individuales
        vcards = re.split(r'END:VCARD\s*', content)

        for vcard in vcards:
            if 'BEGIN:VCARD' not in vcard:
                continue

            contact = self._parse_vcard(vcard)
            if contact:
                self.contacts.append(contact)

        return self.contacts

    def _parse_vcard(self, vcard_text: str) -> Dict:
        """Parsea un vCard individual y extrae todos los campos"""
        contact = {
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

        # Eliminar BEGIN:VCARD si existe
        vcard_text = vcard_text.replace('BEGIN:VCARD', '')

        # Unir líneas continuadas (que empiezan con espacio o tab)
        vcard_text = re.sub(r'\r?\n[ \t]', '', vcard_text)

        lines = vcard_text.strip().split('\n')

        for line in lines:
            line = line.strip()
            if not line or ':' not in line:
                continue

            # Separar campo y valor
            field_part, _, value = line.partition(':')

            # El campo puede tener parámetros como TEL;TYPE=CELL
            field_parts = field_part.split(';')
            field = field_parts[0].upper()
            params = field_parts[1:] if len(field_parts) > 1 else []

            # Extraer tipo (TYPE=CELL, TYPE=HOME, etc.)
            field_type = self._extract_type(params)

            # Procesar según el campo
            if field == 'FN':
                contact['fn'] = self._decode_value(value)

            elif field == 'N':
                # Formato: Apellido;Nombre;Segundo nombre;Prefijo;Sufijo
                parts = value.split(';')
                contact['family_name'] = self._decode_value(parts[0] if len(parts) > 0 else '')
                contact['given_name'] = self._decode_value(parts[1] if len(parts) > 1 else '')
                contact['middle_name'] = self._decode_value(parts[2] if len(parts) > 2 else '')

            elif field == 'TEL':
                phone = self._clean_phone(value)
                if phone:
                    contact['phones'].append({
                        'number': phone,
                        'type': field_type
                    })

            elif field == 'EMAIL':
                email = self._decode_value(value).strip()
                if email and '@' in email:
                    contact['emails'].append({
                        'address': email,
                        'type': field_type
                    })

            elif field == 'NOTE':
                note_text = self._decode_value(value)
                if contact['notes']:
                    contact['notes'] += ' | ' + note_text
                else:
                    contact['notes'] = note_text

            elif field == 'ORG':
                org = self._decode_value(value)
                contact['org'] = org

            elif field == 'ADR':
                # Formato: POBox;Ext;Calle;Ciudad;Estado;CP;País
                addr = self._decode_value(value)
                if addr:
                    contact['addresses'].append({
                        'address': addr.replace(';', ', '),
                        'type': field_type
                    })

            elif field == 'PHOTO':
                # Guardar URL o referencia a foto
                contact['photo'] = value[:100]  # Limitar tamaño

        # Normalizar nombre completo
        contact['fn'] = self._normalize_full_name(contact)

        # Solo devolver si tiene al menos nombre o teléfono
        if contact['fn'] or contact['phones']:
            return contact

        return None

    def _extract_type(self, params: List[str]) -> str:
        """Extrae el tipo de un campo (CELL, HOME, WORK, etc.)"""
        for param in params:
            if param.upper().startswith('TYPE='):
                return param.split('=')[1].capitalize()
            elif param.upper() in ['CELL', 'HOME', 'WORK', 'VOICE', 'FAX', 'PREF', 'MAIN']:
                return param.capitalize()
        return 'Other'

    def _decode_value(self, value: str) -> str:
        """Decodifica valores que pueden estar en QUOTED-PRINTABLE o con escapes"""
        # Remover comillas si existen
        value = value.strip('"')

        # Decodificar QUOTED-PRINTABLE
        if '=' in value and re.search(r'=[0-9A-F]{2}', value):
            try:
                # Simple decodificación de quoted-printable
                value = re.sub(r'=([0-9A-F]{2})', lambda m: chr(int(m.group(1), 16)), value)
            except:
                pass

        # Decodificar escapes comunes
        value = value.replace('\\n', '\n').replace('\\,', ',').replace('\\;', ';')

        return value.strip()

    def _clean_phone(self, phone: str) -> str:
        """Limpia y normaliza un número de teléfono"""
        # Remover espacios, guiones, paréntesis
        phone = re.sub(r'[^\d+]', '', phone)
        return phone if phone else None

    def _normalize_full_name(self, contact: Dict) -> str:
        """Normaliza el nombre completo del contacto"""
        # Si ya existe FN, usarlo
        if contact['fn']:
            return contact['fn']

        # Construir desde N
        parts = []
        if contact['given_name']:
            parts.append(contact['given_name'])
        if contact['middle_name']:
            parts.append(contact['middle_name'])
        if contact['family_name']:
            parts.append(contact['family_name'])

        if parts:
            return ' '.join(parts)

        # Si no hay nombre, usar el primer teléfono
        if contact['phones']:
            return f"Contacto {contact['phones'][0]['number']}"

        return "Sin nombre"


class ContactMerger:
    """Unifica contactos duplicados basándose en nombre o teléfono"""

    def __init__(self, contacts: List[Dict]):
        self.contacts = contacts

    def merge_duplicates(self) -> List[Dict]:
        """Fusiona contactos duplicados"""
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
        """Fusiona un grupo de contactos en uno solo"""
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
    """Genera CSV compatible con Google Contacts"""

    def __init__(self, contacts: List[Dict], output_file: str):
        self.contacts = contacts
        self.output_file = output_file

    def generate(self):
        """Genera el archivo CSV"""
        # Determinar número máximo de teléfonos y emails
        max_phones = max((len(c['phones']) for c in self.contacts), default=0)
        max_emails = max((len(c['emails']) for c in self.contacts), default=0)

        # Limitar a un máximo razonable
        max_phones = min(max_phones, 5)
        max_emails = min(max_emails, 5)

        # Construir encabezados
        headers = [
            'Name',
            'Given Name',
            'Family Name',
        ]

        # Agregar columnas de teléfonos
        for i in range(1, max_phones + 1):
            headers.append(f'Phone {i} - Type')
            headers.append(f'Phone {i} - Value')

        # Agregar columnas de emails
        for i in range(1, max_emails + 1):
            headers.append(f'E-mail {i} - Type')
            headers.append(f'E-mail {i} - Value')

        headers.extend([
            'Organization 1 - Name',
            'Address 1 - Formatted',
            'Notes'
        ])

        # Escribir CSV
        with open(self.output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()

            for contact in self.contacts:
                row = {
                    'Name': contact['fn'],
                    'Given Name': contact['given_name'],
                    'Family Name': contact['family_name'],
                    'Organization 1 - Name': contact['org'],
                    'Notes': contact['notes']
                }

                # Agregar teléfonos
                for i, phone in enumerate(contact['phones'][:max_phones], 1):
                    row[f'Phone {i} - Type'] = phone['type']
                    row[f'Phone {i} - Value'] = phone['number']

                # Agregar emails
                for i, email in enumerate(contact['emails'][:max_emails], 1):
                    row[f'E-mail {i} - Type'] = email['type']
                    row[f'E-mail {i} - Value'] = email['address']

                # Agregar primera dirección
                if contact['addresses']:
                    row['Address 1 - Formatted'] = contact['addresses'][0]['address']

                writer.writerow(row)

        print(f"✓ CSV generado exitosamente: {self.output_file}")


def main():
    """Función principal"""
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

    print(f"\n📂 Archivo de entrada: {input_file}")
    print(f"📄 Archivo de salida: {output_file}\n")

    # 1. Parsear el archivo VCF
    print("⏳ Paso 1/3: Parseando archivo .vcf...")
    parser = VCardParser(input_file)
    contacts = parser.parse()
    print(f"   ✓ {len(contacts)} contactos extraídos")

    # 2. Fusionar duplicados
    print("\n⏳ Paso 2/3: Fusionando contactos duplicados...")
    merger = ContactMerger(contacts)
    merged_contacts = merger.merge_duplicates()
    duplicates_removed = len(contacts) - len(merged_contacts)
    print(f"   ✓ {duplicates_removed} duplicados fusionados")
    print(f"   ✓ {len(merged_contacts)} contactos únicos")

    # 3. Generar CSV
    print("\n⏳ Paso 3/3: Generando CSV compatible con Google Contacts...")
    csv_generator = GoogleContactsCSV(merged_contacts, output_file)
    csv_generator.generate()

    # Estadísticas finales
    print("\n" + "=" * 70)
    print("📊 ESTADÍSTICAS:")
    print(f"   • Contactos procesados: {len(contacts)}")
    print(f"   • Duplicados fusionados: {duplicates_removed}")
    print(f"   • Contactos en CSV final: {len(merged_contacts)}")

    total_phones = sum(len(c['phones']) for c in merged_contacts)
    total_emails = sum(len(c['emails']) for c in merged_contacts)
    print(f"   • Total teléfonos: {total_phones}")
    print(f"   • Total emails: {total_emails}")

    print("\n✅ ¡Proceso completado exitosamente!")
    print(f"   Puedes importar {output_file} directamente en Google Contacts")
    print("=" * 70)


if __name__ == '__main__':
    main()
