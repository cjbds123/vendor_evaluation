from datetime import datetime, timezone
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


# ---------------------------------------------------------------------------
# Scoring Scale  (configurable per-org, seed with CPMS defaults)
# ---------------------------------------------------------------------------
class ScoringLevel(db.Model):
    __tablename__ = 'scoring_level'
    id          = db.Column(db.Integer, primary_key=True)
    label       = db.Column(db.String(80), nullable=False)
    value       = db.Column(db.Integer, nullable=True)  # None = N/A
    sort_order  = db.Column(db.Integer, default=0)
    is_default  = db.Column(db.Boolean, default=False)

    def __repr__(self):
        return f'<ScoringLevel {self.label}={self.value}>'


# ---------------------------------------------------------------------------
# Area  (top-level grouping above Category)
# ---------------------------------------------------------------------------
class Area(db.Model):
    __tablename__ = 'area'
    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(200), nullable=False)
    suite_type  = db.Column(db.String(30), nullable=False, default='Functional')
    test_suite_id = db.Column(db.Integer, db.ForeignKey('test_suite.id'), nullable=True)
    sort_order  = db.Column(db.Integer, default=0)
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    categories  = db.relationship('Category', backref='area', lazy='dynamic',
                                   order_by='Category.sort_order')

    def __repr__(self):
        return f'<Area {self.name}>'


# ---------------------------------------------------------------------------
# Association table: TestSuite ↔ TestCase  (many-to-many)
# ---------------------------------------------------------------------------
suite_tests = db.Table(
    'suite_tests',
    db.Column('suite_id',     db.Integer, db.ForeignKey('test_suite.id',  ondelete='CASCADE'), primary_key=True),
    db.Column('test_case_id', db.Integer, db.ForeignKey('test_case.id',   ondelete='CASCADE'), primary_key=True),
)


# ---------------------------------------------------------------------------
# Test Suite  (a named, per-project collection of test cases)
# ---------------------------------------------------------------------------
class TestSuite(db.Model):
    __tablename__ = 'test_suite'
    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    # Many-to-many with TestCase; 'suites' backref added on TestCase
    test_cases  = db.relationship(
        'TestCase', secondary=suite_tests,
        backref=db.backref('suites', lazy='dynamic'),
        lazy='dynamic',
    )
    # One-to-many reverse: all projects using this suite
    projects    = db.relationship('Project', backref='test_suite', lazy='dynamic')
    # One-to-many: areas belonging to this suite template
    areas       = db.relationship('Area', backref='test_suite', lazy='dynamic',
                                  order_by='Area.sort_order')

    @property
    def test_count(self):
        return self.test_cases.count()

    def __repr__(self):
        return f'<TestSuite {self.name}>'


# ---------------------------------------------------------------------------
# Category  (hierarchical: parent_id for subcategories, belongs to Area)
# ---------------------------------------------------------------------------
class Category(db.Model):
    __tablename__ = 'category'
    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(200), nullable=False)
    suite_type  = db.Column(db.String(30), nullable=False, default='Functional')  # Functional / Non-Functional
    area_id     = db.Column(db.Integer, db.ForeignKey('area.id'), nullable=True)
    parent_id   = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=True)
    sort_order  = db.Column(db.Integer, default=0)
    weight_multiplier = db.Column(db.Float, default=1.0)
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    parent      = db.relationship('Category', remote_side=[id],
                                   backref=db.backref('children', lazy='dynamic',
                                                      order_by='Category.sort_order'))
    test_cases  = db.relationship('TestCase', backref='category', lazy='dynamic')

    def __repr__(self):
        return f'<Category {self.name}>'


# ---------------------------------------------------------------------------
# Test Case
# ---------------------------------------------------------------------------
class TestCase(db.Model):
    __tablename__ = 'test_case'
    id              = db.Column(db.Integer, primary_key=True)
    test_id_code    = db.Column(db.String(20), nullable=False, unique=True)  # e.g. F-001, NF-001
    tier            = db.Column(db.String(20), default='Core')               # Core / Extended
    category_id     = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=False)
    subcategory     = db.Column(db.String(200), nullable=True)
    capability      = db.Column(db.Text, nullable=False)
    test_scenario   = db.Column(db.Text, nullable=True)
    pass_criteria   = db.Column(db.Text, nullable=True)
    evidence_required = db.Column(db.Text, nullable=True)
    test_method     = db.Column(db.String(60), nullable=True)                # Demo, Docs, POC, etc.
    priority        = db.Column(db.String(20), default='Should')             # Must / Should / Could
    weight          = db.Column(db.Float, default=1.0)
    is_mandatory    = db.Column(db.Boolean, default=False)                   # gating requirement
    suite_type      = db.Column(db.String(30), nullable=False, default='Functional')
    sort_order      = db.Column(db.Integer, default=0)
    created_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    results         = db.relationship('TestResult', backref='test_case', lazy='dynamic')

    def __repr__(self):
        return f'<TestCase {self.test_id_code}>'


# ---------------------------------------------------------------------------
# Evaluation Project
# ---------------------------------------------------------------------------
class Project(db.Model):
    __tablename__ = 'project'
    id              = db.Column(db.Integer, primary_key=True)
    name            = db.Column(db.String(200), nullable=False)
    description     = db.Column(db.Text, nullable=True)
    status          = db.Column(db.String(30), default='Active')  # Active / Archived
    test_suite_id   = db.Column(db.Integer, db.ForeignKey('test_suite.id'), nullable=True)
    created_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    vendors         = db.relationship('Vendor', backref='project', lazy='dynamic')

    def __repr__(self):
        return f'<Project {self.name}>'


# ---------------------------------------------------------------------------
# Vendor (per project)
# ---------------------------------------------------------------------------
class Vendor(db.Model):
    __tablename__ = 'vendor'
    id          = db.Column(db.Integer, primary_key=True)
    project_id  = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    name        = db.Column(db.String(200), nullable=False)
    contact     = db.Column(db.String(200), nullable=True)
    notes       = db.Column(db.Text, nullable=True)
    eval_method = db.Column(db.String(60), default='Demo')     # Demo / Sandbox / Pilot
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    results     = db.relationship('TestResult', backref='vendor', lazy='dynamic')

    def __repr__(self):
        return f'<Vendor {self.name}>'


# ---------------------------------------------------------------------------
# Test Result  (one per test_case × vendor)
# ---------------------------------------------------------------------------
class TestResult(db.Model):
    __tablename__ = 'test_result'
    id              = db.Column(db.Integer, primary_key=True)
    vendor_id       = db.Column(db.Integer, db.ForeignKey('vendor.id'), nullable=False)
    test_case_id    = db.Column(db.Integer, db.ForeignKey('test_case.id'), nullable=False)
    support_level   = db.Column(db.String(60), nullable=True)   # matches ScoringLevel.label
    score           = db.Column(db.Integer, nullable=True)
    weighted_score  = db.Column(db.Float, nullable=True)
    status          = db.Column(db.String(30), default='Not Started')  # Not Started / In Progress / Blocked / Submitted / Approved
    pass_fail       = db.Column(db.String(10), nullable=True)          # Pass / Fail / N/A
    notes           = db.Column(db.Text, nullable=True)
    block_reason    = db.Column(db.Text, nullable=True)
    updated_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))
    updated_by      = db.Column(db.String(100), nullable=True)

    evidences       = db.relationship('Evidence', backref='test_result', lazy='dynamic')
    audit_entries   = db.relationship('AuditLog', backref='test_result', lazy='dynamic')

    __table_args__ = (db.UniqueConstraint('vendor_id', 'test_case_id', name='uq_vendor_test'),)

    def __repr__(self):
        return f'<TestResult vendor={self.vendor_id} test={self.test_case_id}>'


# ---------------------------------------------------------------------------
# Evidence
# ---------------------------------------------------------------------------
class Evidence(db.Model):
    __tablename__ = 'evidence'
    id              = db.Column(db.Integer, primary_key=True)
    test_result_id  = db.Column(db.Integer, db.ForeignKey('test_result.id'), nullable=False)
    evidence_type   = db.Column(db.String(20), default='file')   # file / link / text
    filename        = db.Column(db.String(300), nullable=True)
    filepath        = db.Column(db.String(500), nullable=True)
    url             = db.Column(db.String(500), nullable=True)
    text_content    = db.Column(db.Text, nullable=True)
    uploaded_at     = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    uploaded_by     = db.Column(db.String(100), nullable=True)

    def __repr__(self):
        return f'<Evidence {self.evidence_type} for result={self.test_result_id}>'


# ---------------------------------------------------------------------------
# Vendor Question  (questions for vendor, with optional test case link)
# ---------------------------------------------------------------------------
class VendorQuestion(db.Model):
    __tablename__ = 'vendor_question'
    id              = db.Column(db.Integer, primary_key=True)
    project_id      = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False)
    vendor_id       = db.Column(db.Integer, db.ForeignKey('vendor.id'), nullable=True)
    test_result_id  = db.Column(db.Integer, db.ForeignKey('test_result.id'), nullable=True)
    area_id         = db.Column(db.Integer, db.ForeignKey('area.id'), nullable=True)
    category_id     = db.Column(db.Integer, db.ForeignKey('category.id'), nullable=True)
    question_text   = db.Column(db.Text, nullable=False)
    vendor_response = db.Column(db.Text, nullable=True)
    created_at      = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    responded_at    = db.Column(db.DateTime, nullable=True)
    created_by      = db.Column(db.String(100), default='evaluator')
    status          = db.Column(db.String(30), default='Open')  # Open / Answered / Closed

    project         = db.relationship('Project', backref=db.backref('questions', lazy='dynamic'))
    vendor          = db.relationship('Vendor', backref=db.backref('questions', lazy='dynamic'))
    test_result     = db.relationship('TestResult', backref=db.backref('questions', lazy='dynamic'))
    area            = db.relationship('Area', backref=db.backref('questions', lazy='dynamic'))
    category        = db.relationship('Category', backref=db.backref('questions', lazy='dynamic'))

    def __repr__(self):
        return f'<VendorQuestion {self.id} project={self.project_id}>'


# ---------------------------------------------------------------------------
# Audit Log
# ---------------------------------------------------------------------------
class AuditLog(db.Model):
    __tablename__ = 'audit_log'
    id              = db.Column(db.Integer, primary_key=True)
    test_result_id  = db.Column(db.Integer, db.ForeignKey('test_result.id'), nullable=True)
    entity_type     = db.Column(db.String(60), nullable=True)    # TestResult, Category, TestCase, etc.
    entity_id       = db.Column(db.Integer, nullable=True)
    action          = db.Column(db.String(60), nullable=False)   # created, updated, scored, approved
    field_changed   = db.Column(db.String(100), nullable=True)
    old_value       = db.Column(db.Text, nullable=True)
    new_value       = db.Column(db.Text, nullable=True)
    user            = db.Column(db.String(100), default='system')
    timestamp       = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def __repr__(self):
        return f'<AuditLog {self.action} on {self.entity_type}#{self.entity_id}>'


# ---------------------------------------------------------------------------
# Vendor Comment  (meeting notes, observations, general comments per vendor)
# ---------------------------------------------------------------------------
class VendorComment(db.Model):
    __tablename__ = 'vendor_comment'
    id          = db.Column(db.Integer, primary_key=True)
    vendor_id   = db.Column(db.Integer, db.ForeignKey('vendor.id', ondelete='CASCADE'), nullable=False)
    title       = db.Column(db.String(300), nullable=True)        # e.g. "Meeting 2026-03-05"
    body        = db.Column(db.Text, nullable=False)
    created_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc),
                            onupdate=lambda: datetime.now(timezone.utc))
    created_by  = db.Column(db.String(100), default='local_user')

    vendor      = db.relationship('Vendor', backref=db.backref('comments', lazy='dynamic',
                                  order_by='VendorComment.created_at.desc()'))

    def __repr__(self):
        return f'<VendorComment {self.id} vendor={self.vendor_id}>'


# ---------------------------------------------------------------------------
# Vendor Document  (files uploaded at the vendor level – meeting docs, etc.)
# ---------------------------------------------------------------------------
class VendorDocument(db.Model):
    __tablename__ = 'vendor_document'
    id          = db.Column(db.Integer, primary_key=True)
    vendor_id   = db.Column(db.Integer, db.ForeignKey('vendor.id', ondelete='CASCADE'), nullable=False)
    doc_type    = db.Column(db.String(20), default='file')       # file / link
    filename    = db.Column(db.String(300), nullable=True)
    filepath    = db.Column(db.String(500), nullable=True)
    url         = db.Column(db.String(500), nullable=True)
    description = db.Column(db.String(500), nullable=True)
    uploaded_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    uploaded_by = db.Column(db.String(100), default='local_user')

    vendor      = db.relationship('Vendor', backref=db.backref('documents', lazy='dynamic',
                                  order_by='VendorDocument.uploaded_at.desc()'))

    def __repr__(self):
        return f'<VendorDocument {self.doc_type} vendor={self.vendor_id}>'
