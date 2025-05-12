import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import json
import csv
import datetime
import smtplib
from email.mime.text import MIMEText
from cryptography.fernet import Fernet
import base64

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio Automatizado de Emails")
        root.iconbitmap("icone-email.ico")
        self.root.geometry("1000x700")
        
        # Configurações iniciais
        self.current_login = None
        self.setup_directories()
        self.setup_encryption()
        
        # Interface principal
        self.create_notebook()
        self.setup_contacts_tab()
        self.setup_messages_tab()
        self.setup_logins_tab()
        self.setup_review_tab()
        
        # Carregar dados
        self.load_contacts()
        self.load_messages()
        self.load_logins()
    
    def setup_directories(self):
        """Cria os diretórios necessários para o aplicativo"""
        os.makedirs('mensagens', exist_ok=True)
        os.makedirs('logs', exist_ok=True)
    
    def setup_encryption(self):
        """Configura o sistema de criptografia para senhas"""
        self.key = self.get_or_create_key()
    
    def get_or_create_key(self):
        """Obtém ou cria uma chave de criptografia"""
        key_file = 'email_app.key'
        if os.path.exists(key_file):
            with open(key_file, 'rb') as f:
                return f.read()
        key = Fernet.generate_key()
        with open(key_file, 'wb') as f:
            f.write(key)
        return key
    
    def encrypt_data(self, data):
        """Criptografa dados sensíveis"""
        cipher_suite = Fernet(self.key)
        return cipher_suite.encrypt(data.encode()).decode()
    
    def decrypt_data(self, encrypted_data):
        """Descriptografa dados"""
        cipher_suite = Fernet(self.key)
        return cipher_suite.decrypt(encrypted_data.encode()).decode()
    
    def create_notebook(self):
        """Cria o notebook (abas) principal"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)
        
        # Criação das abas
        self.contacts_tab = ttk.Frame(self.notebook)
        self.messages_tab = ttk.Frame(self.notebook)
        self.logins_tab = ttk.Frame(self.notebook)
        self.review_tab = ttk.Frame(self.notebook)
        self.logs_tab = None
        
        self.notebook.add(self.contacts_tab, text="Contatos")
        self.notebook.add(self.messages_tab, text="Mensagens")
        self.notebook.add(self.logins_tab, text="Gerenciar Logins")
        self.notebook.add(self.review_tab, text="Revisão e Envio")
    
    # [SECTION] CONTATOS TAB
    def setup_contacts_tab(self):
        """Configura a aba de contatos"""
        # Frame principal
        main_frame = ttk.Frame(self.contacts_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview para contatos
        self.contacts_tree = ttk.Treeview(main_frame, columns=('Nome', 'Email'), show='headings')
        self.contacts_tree.heading('Nome', text='Nome')
        self.contacts_tree.heading('Email', text='Email')
        self.contacts_tree.column('Nome', width=200)
        self.contacts_tree.column('Email', width=300)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=self.contacts_tree.yview)
        self.contacts_tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        self.contacts_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botões
        buttons_frame = ttk.Frame(self.contacts_tab)
        buttons_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Adicionar", command=self.add_contact).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Editar", command=self.edit_contact).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Excluir", command=self.delete_contact).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Atualizar", command=self.load_contacts).pack(side='right', padx=5)
    
    def load_contacts(self):
        """Carrega contatos do arquivo XML"""
        self.contacts_tree.delete(*self.contacts_tree.get_children())
        
        if not os.path.exists('contatos.xml'):
            return
            
        tree = ET.parse('contatos.xml')
        for contact in tree.getroot().findall('contato'):
            name = contact.find('nome').text
            email = contact.find('email').text
            self.contacts_tree.insert('', 'end', values=(name, email))
    
    def save_contacts(self):
        """Salva contatos no arquivo XML"""
        contacts = []
        for item in self.contacts_tree.get_children():
            name, email = self.contacts_tree.item(item)['values']
            contacts.append((name, email))
        
        root = ET.Element('contatos')
        for name, email in contacts:
            contact = ET.SubElement(root, 'contato')
            ET.SubElement(contact, 'nome').text = name
            ET.SubElement(contact, 'email').text = email
        
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")
        with open('contatos.xml', 'w', encoding='utf-8') as f:
            f.write(xml_str)
    
    def add_contact(self):
        """Abre diálogo para adicionar novo contato"""
        self.contact_dialog()
    
    def edit_contact(self):
        """Abre diálogo para editar contato existente"""
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um contato para editar.")
            return
            
        item = self.contacts_tree.item(selected[0])
        self.contact_dialog(item['values'][0], item['values'][1])
    
    def delete_contact(self):
        """Remove contato selecionado"""
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um contato para excluir.")
            return
            
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir este contato?"):
            self.contacts_tree.delete(selected[0])
            self.save_contacts()
    
    def contact_dialog(self, name="", email=""):
        """Diálogo para adicionar/editar contatos"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Editar Contato" if name else "Novo Contato")
        dialog.resizable(False, False)
        dialog.iconbitmap("icone-email.ico")
        
        # Campos do formulário
        ttk.Label(dialog, text="Nome:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        name_entry.insert(0, name)
        
        ttk.Label(dialog, text="Email:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        email_entry = ttk.Entry(dialog, width=30)
        email_entry.grid(row=1, column=1, padx=5, pady=5)
        email_entry.insert(0, email)
        
        def save():
            new_name = name_entry.get().strip()
            new_email = email_entry.get().strip()
            
            if not new_name or not new_email:
                messagebox.showwarning("Aviso", "Preencha todos os campos.")
                return
                
            # Atualizar ou adicionar
            selected = self.contacts_tree.selection()
            if selected:
                self.contacts_tree.item(selected[0], values=(new_name, new_email))
            else:
                self.contacts_tree.insert('', 'end', values=(new_name, new_email))
            
            self.save_contacts()
            dialog.destroy()
        
        ttk.Button(dialog, text="Salvar", command=save).grid(row=2, column=1, sticky='e', padx=5, pady=10)
    
    # [SECTION] MENSAGENS TAB
    def setup_messages_tab(self):
        """Configura a aba de mensagens"""
        # Frame principal
        main_frame = ttk.Frame(self.messages_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview para mensagens
        self.messages_tree = ttk.Treeview(main_frame, columns=('Assunto', 'Arquivo'), show='headings')
        self.messages_tree.heading('Assunto', text='Assunto')
        self.messages_tree.heading('Arquivo', text='Arquivo')
        self.messages_tree.column('Assunto', width=300)
        self.messages_tree.column('Arquivo', width=200)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=self.messages_tree.yview)
        self.messages_tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        self.messages_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botões
        buttons_frame = ttk.Frame(self.messages_tab)
        buttons_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Adicionar", command=self.add_message).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Editar", command=self.edit_message).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Excluir", command=self.delete_message).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Atualizar", command=self.load_messages).pack(side='right', padx=5)
        
        # Pré-visualização
        ttk.Label(self.messages_tab, text="Pré-visualização:").pack(pady=(10, 0))
        self.message_preview = tk.Text(self.messages_tab, height=10, wrap='word')
        self.message_preview.pack(fill='x', padx=10, pady=(0, 10))
        
        self.messages_tree.bind('<<TreeviewSelect>>', lambda e: self.show_message_preview())
    
    def load_messages(self):
        """Carrega mensagens da pasta mensagens/"""
        self.messages_tree.delete(*self.messages_tree.get_children())
        
        for filename in os.listdir('mensagens'):
            if filename.endswith('.txt'):
                with open(os.path.join('mensagens', filename), 'r', encoding='utf-8') as f:
                    subject = f.readline().strip()
                    self.messages_tree.insert('', 'end', values=(subject, filename))
    
    def show_message_preview(self):
        """Exibe pré-visualização da mensagem selecionada"""
        selected = self.messages_tree.selection()
        if not selected:
            return
            
        filename = self.messages_tree.item(selected[0])['values'][1]
        self.message_preview.delete(1.0, tk.END)
        
        try:
            with open(os.path.join('mensagens', filename), 'r', encoding='utf-8') as f:
                self.message_preview.insert(tk.END, f.read())
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{str(e)}")
    
    def add_message(self):
        """Abre diálogo para nova mensagem"""
        self.message_dialog()
    
    def edit_message(self):
        """Abre diálogo para editar mensagem existente"""
        selected = self.messages_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma mensagem para editar.")
            return
            
        filename = self.messages_tree.item(selected[0])['values'][1]
        try:
            with open(os.path.join('mensagens', filename), 'r', encoding='utf-8') as f:
                content = f.read()
            self.message_dialog(content, filename)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{str(e)}")
    
    def delete_message(self):
        """Remove mensagem selecionada"""
        selected = self.messages_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma mensagem para excluir.")
            return
            
        filename = self.messages_tree.item(selected[0])['values'][1]
        if messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir esta mensagem?"):
            try:
                os.remove(os.path.join('mensagens', filename))
                self.load_messages()
                self.message_preview.delete(1.0, tk.END)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível excluir o arquivo:\n{str(e)}")
    
    def message_dialog(self, content="", filename=""):
        """Diálogo para edição de mensagens"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Editar Mensagem" if filename else "Nova Mensagem")
        dialog.geometry("700x600")
        dialog.iconbitmap("icone-email.ico")
        
        # Campos do formulário
        ttk.Label(dialog, text="Assunto:").pack(pady=(10, 5))
        subject_entry = ttk.Entry(dialog)
        subject_entry.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(dialog, text="Mensagem:").pack(pady=(10, 5))
        message_text = tk.Text(dialog, wrap='word')
        message_text.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Preencher campos se estiver editando
        if content:
            lines = content.split('\n')
            subject_entry.insert(0, lines[0])
            message_text.insert(tk.END, '\n'.join(lines[1:]))
        
        def save():
            subject = subject_entry.get().strip()
            message = message_text.get("1.0", tk.END).strip()
            
            if not subject or not message:
                messagebox.showwarning("Aviso", "Preencha todos os campos.")
                return
                
            # Se for nova mensagem, pedir nome do arquivo
            if not filename:
                suggested_name = f"{subject[:50]}.txt".replace('/', '_').replace('\\', '_')
                filepath = filedialog.asksaveasfilename(
                    initialdir='mensagens',
                    initialfile=suggested_name,
                    title="Salvar mensagem como",
                    defaultextension=".txt",
                    filetypes=(("Arquivos de texto", "*.txt"),))
                
                if not filepath:
                    return
                
                filename_to_save = os.path.basename(filepath)
            else:
                filename_to_save = filename
            
            # Salvar arquivo
            try:
                with open(os.path.join('mensagens', filename_to_save), 'w', encoding='utf-8') as f:
                    f.write(f"{subject}\n{message}")
                self.load_messages()
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível salvar:\n{str(e)}")
        
        ttk.Button(dialog, text="Salvar", command=save).pack(pady=10)
    
    # [SECTION] LOGINS TAB
    def setup_logins_tab(self):
        """Configura a aba de gerenciamento de logins"""
        # Frame principal
        main_frame = ttk.Frame(self.logins_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview para logins
        self.logins_tree = ttk.Treeview(main_frame, columns=('Email', 'Servidor'), show='headings')
        self.logins_tree.heading('Email', text='Email')
        self.logins_tree.heading('Servidor', text='Servidor SMTP')
        self.logins_tree.column('Email', width=250)
        self.logins_tree.column('Servidor', width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=self.logins_tree.yview)
        self.logins_tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        self.logins_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botões
        buttons_frame = ttk.Frame(self.logins_tab)
        buttons_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Adicionar", command=self.add_login).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Editar", command=self.edit_login).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Remover", command=self.remove_login).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Usar para Envio", command=self.use_login).pack(side='right', padx=5)
    
    def load_logins(self):
        """Carrega logins do arquivo JSON"""
        self.logins_tree.delete(*self.logins_tree.get_children())
        
        if not os.path.exists('logins.json'):
            return
            
        try:
            with open('logins.json', 'r') as f:
                logins = json.load(f)
                for login in logins:
                    self.logins_tree.insert('', 'end', values=(login['email'], login['server']))
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar logins:\n{str(e)}")
    
    def save_logins(self, logins):
        """Salva logins no arquivo JSON"""
        try:
            with open('logins.json', 'w') as f:
                json.dump(logins, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar logins:\n{str(e)}")
            return False
    
    def add_login(self):
        """Abre diálogo para novo login"""
        self.login_dialog()
    
    def edit_login(self):
        """Abre diálogo para editar login existente"""
        selected = self.logins_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um login para editar.")
            return
            
        email = self.logins_tree.item(selected[0])['values'][0]
        
        try:
            with open('logins.json', 'r') as f:
                logins = json.load(f)
                login = next((x for x in logins if x['email'] == email), None)
            
            if login:
                self.login_dialog(login)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar login:\n{str(e)}")
    
    def remove_login(self):
        """Remove login selecionado"""
        selected = self.logins_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um login para remover.")
            return
            
        email = self.logins_tree.item(selected[0])['values'][0]
        
        if messagebox.askyesno("Confirmar", f"Remover o login {email}?"):
            try:
                with open('logins.json', 'r') as f:
                    logins = json.load(f)
                
                logins = [x for x in logins if x['email'] != email]
                
                if self.save_logins(logins):
                    self.load_logins()
                    if self.current_login and self.current_login['email'] == email:
                        self.current_login = None
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível remover login:\n{str(e)}")
    
    def use_login(self):
        """Define login selecionado para uso no envio"""
        selected = self.logins_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um login para usar.")
            return
            
        email = self.logins_tree.item(selected[0])['values'][0]
        
        try:
            with open('logins.json', 'r') as f:
                logins = json.load(f)
                login = next((x for x in logins if x['email'] == email), None)
            
            if login:
                self.current_login = {
                    'email': login['email'],
                    'password': self.decrypt_data(login['password']),
                    'server': login['server'],
                    'port': login.get('port', 587)
                }
                messagebox.showinfo("Sucesso", f"Login {email} selecionado para envio!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar login:\n{str(e)}")
    
    def login_dialog(self, login=None):
        """Diálogo para adicionar/editar logins"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Editar Login" if login else "Novo Login")
        dialog.resizable(False, False)
        dialog.iconbitmap("icone-email.ico")
        
        # Campos do formulário
        ttk.Label(dialog, text="Email:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        email_entry = ttk.Entry(dialog, width=30)
        email_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Senha:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        password_entry = ttk.Entry(dialog, width=30, show='*')
        password_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Servidor SMTP:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        server_entry = ttk.Entry(dialog, width=30)
        server_entry.grid(row=2, column=1, padx=5, pady=5)
        server_entry.insert(0, "smtp.gmail.com")
        
        ttk.Label(dialog, text="Porta:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        port_entry = ttk.Entry(dialog, width=30)
        port_entry.grid(row=3, column=1, padx=5, pady=5)
        port_entry.insert(0, "587")
        
        # Preencher campos se estiver editando
        if login:
            email_entry.insert(0, login['email'])
            password_entry.insert(0, self.decrypt_data(login['password']))
            server_entry.insert(0, login['server'])
            port_entry.insert(0, str(login.get('port', 587)))
        
        def save():
            email = email_entry.get().strip()
            password = password_entry.get().strip()
            server = server_entry.get().strip()
            port = port_entry.get().strip()
            
            if not all([email, password, server, port]):
                messagebox.showwarning("Aviso", "Preencha todos os campos.")
                return
                
            try:
                port = int(port)
            except ValueError:
                messagebox.showwarning("Aviso", "Porta deve ser um número.")
                return
            
            # Carregar logins existentes
            logins = []
            if os.path.exists('logins.json'):
                try:
                    with open('logins.json', 'r') as f:
                        logins = json.load(f)
                except:
                    pass
            
            # Criptografar senha
            encrypted_password = self.encrypt_data(password)
            
            # Atualizar ou adicionar login
            existing = next((i for i, x in enumerate(logins) if x['email'] == email), None)
            new_login = {
                'email': email,
                'password': encrypted_password,
                'server': server,
                'port': port
            }
            
            if existing is not None:
                logins[existing] = new_login
            else:
                logins.append(new_login)
            
            if self.save_logins(logins):
                self.load_logins()
                dialog.destroy()
        
        ttk.Button(dialog, text="Salvar", command=save).grid(row=4, column=1, sticky='e', padx=5, pady=10)
    
    # [SECTION] REVIEW TAB
    def setup_review_tab(self):
        """Configura a aba de revisão e envio"""
        # Frame principal
        main_frame = ttk.Frame(self.review_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Frame para contatos selecionados
        contacts_frame = ttk.LabelFrame(main_frame, text="Contatos Selecionados")
        contacts_frame.pack(fill='x', pady=5)
        
        self.selected_contacts_listbox = tk.Listbox(contacts_frame, height=5)
        self.selected_contacts_listbox.pack(fill='x', expand=True)
        
        # Frame para mensagem selecionada
        message_frame = ttk.LabelFrame(main_frame, text="Mensagem Selecionada")
        message_frame.pack(fill='both', expand=True, pady=5)
        
        self.selected_message_text = tk.Text(message_frame, height=15, wrap='word')
        self.selected_message_text.pack(fill='both', expand=True)
        
        # Frame para botões
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=5)
        
        ttk.Button(buttons_frame, text="Selecionar Contatos", command=self.select_contacts).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Selecionar Mensagem", command=self.select_message).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Enviar Emails", command=self.send_emails).pack(side='right', padx=5)
    
    def select_contacts(self):
        """Seleciona contatos para envio"""
        selected_items = self.contacts_tree.selection()
        self.selected_contacts_listbox.delete(0, tk.END)
        
        for item in selected_items:
            name, email = self.contacts_tree.item(item)['values']
            self.selected_contacts_listbox.insert(tk.END, f"{name} <{email}>")
    
    def select_message(self):
        """Seleciona mensagem para envio"""
        selected = self.messages_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione uma mensagem para enviar.")
            return
            
        filename = self.messages_tree.item(selected[0])['values'][1]
        try:
            with open(os.path.join('mensagens', filename), 'r', encoding='utf-8') as f:
                self.selected_message_text.delete(1.0, tk.END)
                self.selected_message_text.insert(tk.END, f.read())
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a mensagem:\n{str(e)}")
    
    def send_emails(self):
        """Envia emails para os contatos selecionados"""
        # Verificar condições para envio
        if not self.current_login:
            messagebox.showwarning("Aviso", "Selecione um login na aba 'Gerenciar Logins' primeiro.")
            return
            
        if self.selected_contacts_listbox.size() == 0:
            messagebox.showwarning("Aviso", "Selecione pelo menos um contato.")
            return
            
        message_content = self.selected_message_text.get("1.0", tk.END).strip()
        if not message_content:
            messagebox.showwarning("Aviso", "Selecione uma mensagem para enviar.")
            return
        
        # Extrair assunto e corpo
        subject = message_content.split('\n')[0]
        body = '\n'.join(message_content.split('\n')[1:])
        
        # Configurar servidor SMTP
        try:
            server = smtplib.SMTP(self.current_login['server'], self.current_login['port'])
            server.starttls()
            server.login(self.current_login['email'], self.current_login['password'])
            
            # Enviar para cada contato
            for i in range(self.selected_contacts_listbox.size()):
                contact = self.selected_contacts_listbox.get(i)
                recipient = contact.split('<')[1].split('>')[0].strip()
                
                # Criar mensagem MIME
                msg = MIMEText(body)
                msg['Subject'] = subject
                msg['From'] = self.current_login['email']
                msg['To'] = recipient
                
                try:
                    server.sendmail(self.current_login['email'], recipient, msg.as_string())
                    self.log_email(self.current_login['email'], recipient, subject, "Sucesso")
                except Exception as e:
                    self.log_email(self.current_login['email'], recipient, subject, f"Falha: {str(e)}")
            
            server.quit()
            messagebox.showinfo("Sucesso", f"Emails enviados para {self.selected_contacts_listbox.size()} contatos!")
            self.add_logs_tab()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha no envio:\n{str(e)}")
    
    # [SECTION] LOGS SYSTEM
    def log_email(self, sender, recipient, subject, status):
        """Registra um envio no arquivo de logs"""
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = [timestamp, sender, recipient, subject, status]
        
        try:
            with open('logs/envios.csv', 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(log_entry)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível registrar o envio:\n{str(e)}")
    
    def add_logs_tab(self):
        """Adiciona/atualiza a aba de logs"""
        # Remover aba existente se houver
        for tab_id in self.notebook.tabs():
            if self.notebook.tab(tab_id, 'text') == "Registro de Envios":
                self.notebook.forget(tab_id)
                break
        
        # Criar nova aba
        self.logs_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.logs_tab, text="Registro de Envios")
        self.notebook.select(self.logs_tab)
        
        # Treeview para logs
        logs_tree = ttk.Treeview(self.logs_tab, columns=('Data', 'Remetente', 'Destinatário', 'Assunto', 'Status'), show='headings')
        logs_tree.heading('Data', text='Data/Hora')
        logs_tree.heading('Remetente', text='Remetente')
        logs_tree.heading('Destinatário', text='Destinatário')
        logs_tree.heading('Assunto', text='Assunto')
        logs_tree.heading('Status', text='Status')
        
        # Configurar colunas
        logs_tree.column('Data', width=150)
        logs_tree.column('Remetente', width=150)
        logs_tree.column('Destinatário', width=150)
        logs_tree.column('Assunto', width=200)
        logs_tree.column('Status', width=150)
        
        logs_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Carregar logs
        try:
            with open('logs/envios.csv', 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader)  # Pular cabeçalho
                for row in reader:
                    logs_tree.insert('', 'end', values=row)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar os logs:\n{str(e)}")

# Inicialização do aplicativo
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
