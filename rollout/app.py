from flask import Flask, render_template, request, redirect, url_for, flash, Response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
import csv
from io import StringIO
from datetime import datetime
from flask import send_file
import pandas as pd
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///rollout.db'
app.config['SECRET_KEY'] = 'your_secret_key'  
app.config['UPLOAD_FOLDER'] = 'static/uploads'
db = SQLAlchemy(app)

# Model untuk tabel rollout
class Rollout(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    nik = db.Column(db.String(20), nullable=False)
    jenis_update = db.Column(db.String(50), nullable=False)
    aplikasi = db.Column(db.String(100), nullable=False)
    menu_update = db.Column(db.String(100), nullable=False)
    versi_aplikasi = db.Column(db.String(50), nullable=False)
    surat_pernyataan = db.Column(db.String(100), nullable=False)
    tanggal_awal_rollout = db.Column(db.DateTime, nullable=False)
    tanggal_akhir_rollout = db.Column(db.DateTime, nullable=False)
    keterangan = db.Column(db.String(255))
    lampiran = db.Column(db.String(255), nullable=False)

# Membuat tabel jika belum ada
def create_tables():
    with app.app_context():
        db.create_all()

# Beranda
@app.route('/')
def home():
    return render_template('home.html')

# nambah data
from datetime import datetime

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        name = request.form['name']
        nik = request.form['nik']
        jenis_update = request.form['jenis_update']
        aplikasi = request.form['aplikasi']
        menu_update = request.form['menu_update']
        versi_aplikasi = request.form['versi_aplikasi']
        surat_pernyataan = request.form['surat_pernyataan']
        tanggal_awal_rollout = datetime.fromisoformat(request.form['tanggal_awal_rollout'])
        tanggal_akhir_rollout = datetime.fromisoformat(request.form['tanggal_akhir_rollout'])
        keterangan = request.form['keterangan']
        lampiran = request.files['lampiran']

        lampiran_path = f'static/uploads/{lampiran.filename}'
        lampiran.save(lampiran_path)

        new_rollout = Rollout(
            name=name,
            nik=nik,
            jenis_update=jenis_update,
            aplikasi=aplikasi,
            menu_update=menu_update,
            versi_aplikasi=versi_aplikasi,
            surat_pernyataan=surat_pernyataan,
            tanggal_awal_rollout=tanggal_awal_rollout,
            tanggal_akhir_rollout=tanggal_akhir_rollout,
            keterangan=keterangan,
            lampiran=lampiran_path
        )

        db.session.add(new_rollout)
        db.session.commit()

        return redirect(url_for('view'))
    return render_template('upload.html')

# liat data
@app.route('/view', methods=['GET'])
def view():
    search_query = request.args.get('search')
    if search_query:
        rollouts = Rollout.query.filter(
            (Rollout.name.ilike(f"%{search_query}%")) |
            (Rollout.nik.ilike(f"%{search_query}%")) |
            (Rollout.aplikasi.ilike(f"%{search_query}%"))
        ).order_by(Rollout.id.desc()).all()  # Mengurutkan berdasarkan ID (descending)
    else:
        rollouts = Rollout.query.order_by(Rollout.id.desc()).all()  # Mengurutkan berdasarkan ID (descending)

    return render_template('view.html', rollouts=rollouts, search_query=search_query)


# edit data
@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit(id):
    rollout = Rollout.query.get_or_404(id)
    
    if request.method == 'POST':
        rollout.name = request.form['name']
        rollout.nik = request.form['nik']
        rollout.jenis_update = request.form['jenis_update']
        rollout.aplikasi = request.form['aplikasi']
        rollout.menu_update = request.form['menu_update']
        rollout.versi_aplikasi = request.form['versi_aplikasi']
        rollout.surat_pernyataan = request.form['surat_pernyataan']
        rollout.tanggal_awal_rollout = datetime.strptime(request.form['tanggal_awal_rollout'], '%Y-%m-%d')
        rollout.tanggal_akhir_rollout = datetime.strptime(request.form['tanggal_akhir_rollout'], '%Y-%m-%d')
        rollout.keterangan = request.form['keterangan']

        lampiran = request.files['lampiran']
        if lampiran:
            lampiran_filename = lampiran.filename
            lampiran_path = os.path.join(app.config['UPLOAD_FOLDER'], lampiran_filename)
            lampiran.save(lampiran_path)
            rollout.lampiran = lampiran_filename

        db.session.commit()
        flash('Data berhasil diupdate!', 'success')
        return redirect(url_for('view'))

    return render_template('edit.html', rollout=rollout)

# hapus data
@app.route('/delete/<int:id>')
def delete(id):
    rollout = Rollout.query.get_or_404(id)
    db.session.delete(rollout)
    db.session.commit()
    flash('Data berhasil dihapus!', 'success')
    return redirect(url_for('view'))

#grafik
@app.route('/graph')
def graph():
    # Ambil data database menggunakan strftime untuk mengambil bulan
    data = db.session.query(
        Rollout.aplikasi,
        db.func.strftime('%m', Rollout.tanggal_awal_rollout).label('bulan'),
        db.func.count(Rollout.aplikasi).label('jumlah'),
        Rollout.name
    ).group_by(Rollout.aplikasi, 'bulan', Rollout.name).all()

    # Format grafik
    chart_data = {}
    for row in data:
        month_name = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 
                      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][int(row.bulan) - 1]
        if month_name not in chart_data:
            chart_data[month_name] = []
        chart_data[month_name].append({'aplikasi': row.aplikasi, 'jumlah': row.jumlah, 'name': row.name})

    return render_template('graph.html', chart_data=chart_data)

#download excellll
@app.route('/download', methods=['GET'])
def download():
    # Ambil data dari database
    data = Rollout.query.all()
    
    # Buat DataFrame dari data
    df = pd.DataFrame([{
        "Name": item.name,
        "NIK": item.nik,
        "Jenis Update": item.jenis_update,
        "Aplikasi": item.aplikasi,
        "Menu Update": item.menu_update,
        "Versi Aplikasi": item.versi_aplikasi,
        "Surat Pernyataan": item.surat_pernyataan,
        "Tanggal Awal Rollout": item.tanggal_awal_rollout.strftime('%Y-%m-%d'),
        "Tanggal Akhir Rollout": item.tanggal_akhir_rollout.strftime('%Y-%m-%d'),
        "Keterangan": item.keterangan
    } for item in data])

    # Simpan DataFrame ke dalam buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data Rollout')
    output.seek(0)

    # Kirim file sebagai response
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, download_name='data_rollout.xlsx')

if __name__ == '__main__':
    create_tables()
    app.run(debug=True)
